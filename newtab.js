
// ==================== TYPE ICON HELPER ====================
function getTypeIconHtml(type, participantText) {
  if (!type) return '';
  const isSlackThread = type === 'slack_thread';
  const isSlackMessage = type === 'slack_channel_dm' || type === 'slack_dm' || type === 'slack_message';
  if (!isSlackThread && !isSlackMessage) return ''; // No icon for gmail, drive, etc.
  const color = '#667eea';
  if (isSlackThread) {
    return `<span class="task-type-icon" title="Slack thread" style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;opacity:0.6;flex-shrink:0;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M9 6h10"/><path d="M9 12h10"/><path d="M9 18h10"/><circle cx="4" cy="6" r="1.5" fill="${color}" stroke="none"/><circle cx="4" cy="12" r="1.5" fill="${color}" stroke="none"/><circle cx="4" cy="18" r="1.5" fill="${color}" stroke="none"/></svg></span>`;
  } else {
    const pt = (participantText || '').toLowerCase();
    const isDM = type === 'slack_dm' || (pt && !pt.startsWith('#') && !pt.toLowerCase().includes('channel'));
    const msgTitle = isDM ? 'Message in the DM' : 'Message in the channel';
    return `<span class="task-type-icon" title="${msgTitle}" style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;opacity:0.6;flex-shrink:0;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></span>`;
  }
}

function getTrendingIconHtml(task) {
  if (task.trending !== 'true' && task.trending !== true) return '';
  const updatedAt = new Date(task.updated_at || task.created_at);
  const oneHourAgo = new Date(Date.now() - 60 * 60 * 1000);
  if (updatedAt < oneHourAgo) return '';
  return `<span class="trending-buzz-icon"><svg width="14" height="14" viewBox="0 0 24 24" fill="#2ecc71" xmlns="http://www.w3.org/2000/svg"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 15v-4H7l5-7v4h4l-5 7z"/></svg><span class="trending-tooltip">Active discussion underway. More than 3 messages in last 1 hour.</span></span>`;
}

// Complete newtab.js with Scratchpad, Bookmarks, and Todos functionality
const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';
const AUTH_URL = 'https://n8n-kqq5.onrender.com/webhook/e6bcd2c3-c714-46c7-94b8-8aeb9831429c';
const STORAGE_KEY = 'oracle_user_data';
const READ_TASKS_KEY = 'oracle_read_tasks';

// Helper: Extract Slack channel ID from a message link
function extractSlackChannelId(messageLink) {
  if (!messageLink) return null;
  const match = messageLink.match(/slack\.com\/archives\/([A-Z0-9]+)/i);
  return match ? match[1] : null;
}

function getSlackWorkspaceDomain() {
  // Extract workspace domain (e.g. "fwbuzz.slack.com") from any known Slack message_link
  const allItems = (typeof allTodos !== 'undefined' ? allTodos : [])
    .concat(typeof allFyiItems !== 'undefined' ? allFyiItems : []);
  for (const item of allItems) {
    const link = item.message_link || '';
    const m = link.match(/https?:\/\/([^/]+\.slack\.com)/i);
    if (m) return m[1];
  }
  return 'slack.com'; // fallback
}
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
const { loadReadState, saveReadState: _origSaveReadState, markTaskAsRead, isTaskUnread, isValidTag, sortTodos, formatDate, formatDueBy, escapeHtml } = window.Oracle;
// Keep original saveReadState (badge is updated via updateGlobalUpdatesBadge)
const saveReadState = _origSaveReadState;
const { isMeetingLink, isDriveLink, isSlackLink, getSlackChannelUrl, extractDriveFileId, getCleanDriveFileUrl } = window.OracleIcons;
const { groupTasksByDriveFile, groupTasksBySlackChannel, groupTasksByTag, extractFileNameFromTask } = window.OracleGrouping;

function markTaskAsUnread(todoId) { readTaskIds.delete(String(todoId)); saveReadState(); updateExtensionBadge(); }

// Update Chrome extension badge with pending update counts
function updateExtensionBadge() {
  try {
    const actionPending = (typeof pendingActionUpdates !== 'undefined') ? pendingActionUpdates.length : 0;
    const fyiPending = (typeof pendingFyiUpdates !== 'undefined') ? pendingFyiUpdates.length : 0;
    const total = actionPending + fyiPending;
    if (chrome && chrome.runtime && chrome.runtime.sendMessage) {
      chrome.runtime.sendMessage({
        type: 'updateBadge',
        actionUnread: actionPending,
        fyiUnread: fyiPending,
        total
      }).catch(() => {});
    }
  } catch (e) { /* ignore */ }
}
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
let keyboardSelectedFeedId = null;
let currentKeyboardColumn = null; // 'action', 'fyi', 'feed', or 'notes'

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
  const starredCount = allTodos.filter(t => t.starred === 1 && !isMeetingLink(t.message_link) && !(t.tags || []).some(tag => String(tag) === '2282' || String(tag).toLowerCase() === 'daily feed')).length;
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
  const isSliderOpen = document.querySelector('.transcript-slider-overlay') || document.querySelector('.note-viewer-overlay') || document.querySelector('.chat-slider-overlay');
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
    if (keyboardSelectedTaskId && (currentKeyboardColumn === 'action' || currentKeyboardColumn === 'fyi' || currentKeyboardColumn === 'alltasks')) {
      e.preventDefault();
      e.stopPropagation();
      markKeyboardSelectedTaskDone();
      return;
    }
  }

  // Escape - Clear all selections or close note form or close transcript slider
  if (e.key === 'Escape') {
    // Don't handle Escape if attachment preview modal is open — let its own handler deal with it
    if (document.querySelector('.attachment-preview-modal')) return;
    if (isNoteFormActive) {
      e.preventDefault();
      if (typeof window.hideNoteForm === 'function') {
        window.hideNoteForm();
      }
      return;
    }
    // Close transcript/feed slider if open
    if (isSliderOpen) {
      e.preventDefault();
      const closeBtn = document.querySelector('.transcript-close-btn') || document.querySelector('.feed-oracle-close');
      if (closeBtn) closeBtn.click();
      return;
    }
    if (!isInputFocused) {
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
    window.OracleNewMessage.showNewMessageSlider({ mode: 'col3' });
    return;
  }

  // Press "R" - Trigger refresh (click the "X new updates" badge)
  if (e.key === 'r' || e.key === 'R') {
    e.preventDefault();
    const badge = document.getElementById('globalUpdatesBadge');
    if (badge && badge.style.display !== 'none') {
      badge.click();
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

  // Press "3" - Select first item in Daily Feed
  if (e.key === '3') {
    e.preventDefault();
    selectFirstFeedItem();
    return;
  }

  // Press "4" - Select first note in Notes column
  if (e.key === '4') {
    e.preventDefault();
    selectFirstNoteInColumn();
    return;
  }

  // Press "5" - Select first task in All Tasks section
  if (e.key === '5') {
    e.preventDefault();
    selectFirstAllTasksItem();
    return;
  }

  // Press "C" - Copy message_link of selected task to clipboard (works in action, fyi, feed, all tasks)
  if ((e.key === 'c' || e.key === 'C') && !e.metaKey && !e.ctrlKey) {
    const taskId = keyboardSelectedTaskId || keyboardSelectedFeedId;
    if (taskId && (currentKeyboardColumn === 'action' || currentKeyboardColumn === 'fyi' || currentKeyboardColumn === 'feed' || currentKeyboardColumn === 'alltasks')) {
      e.preventDefault();
      const allItems = [...(allTodos || []), ...(allFyiItems || []), ...(allCalendarItems || []), ...(typeof allCompletedTasks !== 'undefined' ? allCompletedTasks : [])];
      const task = allItems.find(t => String(t.id) === String(taskId));
      if (task && task.message_link) {
        navigator.clipboard.writeText(task.message_link).then(() => {
          // Brief visual feedback on the selected task
          const taskEl = document.querySelector(`.todo-item[data-todo-id="${taskId}"], .dailyfeed-item[data-task-id="${taskId}"], .alltasks-task-item[data-task-id="${taskId}"]`);
          if (taskEl) {
            const origBg = taskEl.style.background;
            taskEl.style.background = 'rgba(102,126,234,0.2)';
            setTimeout(() => { taskEl.style.background = origBg; }, 500);
          }
        });
      }
      return;
    }
  }

  // Feed column shortcuts when feed column is active
  if (currentKeyboardColumn === 'feed') {
    if (e.key === 'Enter' && !e.shiftKey && !e.metaKey && !e.ctrlKey && keyboardSelectedFeedId) {
      e.preventDefault();
      const taskId = parseInt(keyboardSelectedFeedId);
      if (taskId) {
        markTaskAsRead(taskId);
        const item = document.querySelector(`.dailyfeed-item[data-task-id="${keyboardSelectedFeedId}"]`);
        if (item) item.classList.remove('unread');
        if (typeof window.showFeedOracleSlider === 'function') window.showFeedOracleSlider(taskId);
      }
      return;
    }
    // Shift+Enter - Mark selected feed items as done
    if (e.shiftKey && e.key === 'Enter' && keyboardSelectedFeedId) {
      e.preventDefault();
      e.stopPropagation();
      if (selectedTodoIds.size > 0) {
        handleBulkMarkDone();
      } else if (keyboardSelectedFeedId) {
        markFeedItemDone(keyboardSelectedFeedId);
      }
      return;
    }
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (e.shiftKey) {
        extendFeedSelectionDown();
      } else {
        moveFeedSelection('down');
      }
      return;
    }
    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (e.shiftKey) {
        extendFeedSelectionUp();
      } else {
        moveFeedSelection('up');
      }
      return;
    }
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
      const notes = (window.Oracle && window.Oracle.state && window.Oracle.state.allNotes) || allNotes || [];
      const note = notes.find(n => n.id == keyboardSelectedNoteId);
      if (note && typeof window.showNoteEditSlider === 'function') {
        window.showNoteEditSlider(note);
      }
      return;
    }

    // Enter - Open note in editable slider
    if (e.key === 'Enter' && keyboardSelectedNoteId) {
      e.preventDefault();
      const notes = (window.Oracle && window.Oracle.state && window.Oracle.state.allNotes) || allNotes || [];
      const note = notes.find(n => n.id == keyboardSelectedNoteId);
      if (note && typeof window.showNoteEditSlider === 'function') {
        window.showNoteEditSlider(note);
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
  if (keyboardSelectedTaskId && (currentKeyboardColumn === 'action' || currentKeyboardColumn === 'fyi' || currentKeyboardColumn === 'alltasks') && !isInInput) {
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

  // Find the list container
  const listContainer = columnEl.querySelector('.fyi-list, .todos-list');
  if (!listContainer) return;

  // Walk direct children in DOM order to find the first selectable item
  for (const child of listContainer.children) {
    // Skip meetings/documents accordions
    if (child.classList.contains('meetings-accordion') || child.classList.contains('documents-accordion')) continue;

    if (child.classList.contains('task-group')) {
      // First item is a task group — expand and select first task inside
      expandTaskGroupAndSelectFirst(child);
      return;
    }

    if (child.classList.contains('todo-item')) {
      // First item is a standalone task
      keyboardSelectedTaskId = child.dataset.todoId;
      highlightKeyboardSelectedTask();
      child.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      return;
    }
  }
}

function moveKeyboardSelection(direction) {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  // Find current task element
  const currentTask = columnEl.querySelector(`.todo-item[data-todo-id="${keyboardSelectedTaskId}"]`);
  if (!currentTask) return;

  const listContainer = currentTask.closest('.fyi-list, .todos-list') || currentTask.parentElement;
  if (!listContainer) return;

  const currentTopLevel = currentTask.closest('.task-group') || currentTask;

  if (direction === 'down') {
    // Check if inside an expanded group — try next sibling in group first
    const taskGroupTasks = currentTask.closest('.task-group-tasks');
    if (taskGroupTasks) {
      const siblingsInGroup = Array.from(taskGroupTasks.querySelectorAll('.todo-item'));
      const idxInGroup = siblingsInGroup.indexOf(currentTask);
      if (idxInGroup < siblingsInGroup.length - 1) {
        clearKeyboardSelection();
        keyboardSelectedTaskId = siblingsInGroup[idxInGroup + 1].dataset.todoId;
        highlightKeyboardSelectedTask();
        siblingsInGroup[idxInGroup + 1].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        return;
      }
    }

    // Move to next top-level item
    const topLevelItems = Array.from(listContainer.children).filter(el =>
      (el.classList.contains('todo-item') || el.classList.contains('task-group')) &&
      !el.closest('.meetings-accordion') && !el.closest('.documents-accordion')
    );
    const topIndex = topLevelItems.indexOf(currentTopLevel);
    if (topIndex === -1 || topIndex >= topLevelItems.length - 1) return;

    const nextTopLevel = topLevelItems[topIndex + 1];
    if (nextTopLevel.classList.contains('task-group')) {
      // Expand if collapsed, select first item
      const tasksContainer = nextTopLevel.querySelector('.task-group-tasks');
      const chevron = nextTopLevel.querySelector('.task-group-chevron');
      if (tasksContainer && tasksContainer.style.display === 'none') {
        tasksContainer.style.display = 'block';
        if (chevron) chevron.style.transform = 'rotate(180deg)';
        nextTopLevel.classList.add('expanded');
      }
      const firstItem = tasksContainer?.querySelector('.todo-item');
      if (firstItem) {
        clearKeyboardSelection();
        keyboardSelectedTaskId = firstItem.dataset.todoId;
        highlightKeyboardSelectedTask();
        firstItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    } else if (nextTopLevel.classList.contains('todo-item')) {
      clearKeyboardSelection();
      keyboardSelectedTaskId = nextTopLevel.dataset.todoId;
      highlightKeyboardSelectedTask();
      nextTopLevel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }

  } else {
    // direction === 'up'
    const taskGroupTasks = currentTask.closest('.task-group-tasks');
    if (taskGroupTasks) {
      const siblingsInGroup = Array.from(taskGroupTasks.querySelectorAll('.todo-item'));
      const idxInGroup = siblingsInGroup.indexOf(currentTask);
      if (idxInGroup > 0) {
        clearKeyboardSelection();
        keyboardSelectedTaskId = siblingsInGroup[idxInGroup - 1].dataset.todoId;
        highlightKeyboardSelectedTask();
        siblingsInGroup[idxInGroup - 1].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        return;
      }
    }

    // Move to previous top-level item
    const topLevelItems = Array.from(listContainer.children).filter(el =>
      (el.classList.contains('todo-item') || el.classList.contains('task-group')) &&
      !el.closest('.meetings-accordion') && !el.closest('.documents-accordion')
    );
    const topIndex = topLevelItems.indexOf(currentTopLevel);
    if (topIndex <= 0) return;

    const prevTopLevel = topLevelItems[topIndex - 1];
    if (prevTopLevel.classList.contains('task-group')) {
      // Select last item in the group (expand if needed)
      const tasksContainer = prevTopLevel.querySelector('.task-group-tasks');
      const chevron = prevTopLevel.querySelector('.task-group-chevron');
      if (tasksContainer && tasksContainer.style.display === 'none') {
        tasksContainer.style.display = 'block';
        if (chevron) chevron.style.transform = 'rotate(180deg)';
        prevTopLevel.classList.add('expanded');
      }
      const allItems = tasksContainer?.querySelectorAll('.todo-item');
      const lastItem = allItems?.[allItems.length - 1];
      if (lastItem) {
        clearKeyboardSelection();
        keyboardSelectedTaskId = lastItem.dataset.todoId;
        highlightKeyboardSelectedTask();
        lastItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      }
    } else if (prevTopLevel.classList.contains('todo-item')) {
      clearKeyboardSelection();
      keyboardSelectedTaskId = prevTopLevel.dataset.todoId;
      highlightKeyboardSelectedTask();
      prevTopLevel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  }
}

function extendSelectionDown() {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  // Find the current task element
  const currentTask = columnEl.querySelector(`.todo-item[data-todo-id="${keyboardSelectedTaskId}"]`);
  if (!currentTask) return;

  // Find the list container
  const listContainer = currentTask.closest('.fyi-list, .todos-list') || currentTask.parentElement;
  if (!listContainer) return;

  // Determine the "top-level" element for the current task (itself or its parent task-group)
  const currentTopLevel = currentTask.closest('.task-group') || currentTask;

  // If current task is inside an expanded task-group, check if there's a next sibling within the group
  const taskGroupTasks = currentTask.closest('.task-group-tasks');
  if (taskGroupTasks) {
    const siblingsInGroup = Array.from(taskGroupTasks.querySelectorAll('.todo-item'));
    const idxInGroup = siblingsInGroup.indexOf(currentTask);
    if (idxInGroup < siblingsInGroup.length - 1) {
      // Move to next item within the same group
      const nextInGroup = siblingsInGroup[idxInGroup + 1];
      keyboardSelectedTaskId = nextInGroup.dataset.todoId;
      if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
        toggleTodoSelection(keyboardSelectedTaskId);
      }
      highlightKeyboardSelectedTask();
      nextInGroup.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      return;
    }
    // Otherwise fall through to look at next top-level sibling after this group
  }

  // Get all top-level children of the list (standalone .todo-item and .task-group)
  const topLevelItems = Array.from(listContainer.children).filter(el =>
    (el.classList.contains('todo-item') || el.classList.contains('task-group')) &&
    !el.closest('.meetings-accordion') &&
    !el.closest('.documents-accordion')
  );

  const topIndex = topLevelItems.indexOf(currentTopLevel);
  if (topIndex === -1 || topIndex >= topLevelItems.length - 1) return;

  // Look at the next top-level item
  const nextTopLevel = topLevelItems[topIndex + 1];

  if (nextTopLevel.classList.contains('task-group')) {
    // It's a task group — expand if collapsed, then select first item
    expandTaskGroupAndSelectFirst(nextTopLevel, true);
  } else if (nextTopLevel.classList.contains('todo-item')) {
    // It's a standalone task — select it
    keyboardSelectedTaskId = nextTopLevel.dataset.todoId;
    if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
      toggleTodoSelection(keyboardSelectedTaskId);
    }
    highlightKeyboardSelectedTask();
    nextTopLevel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function expandTaskGroupAndSelectFirst(groupEl, multiSelect = false) {
  const tasksContainer = groupEl.querySelector('.task-group-tasks');
  const chevron = groupEl.querySelector('.task-group-chevron');
  if (tasksContainer && tasksContainer.style.display === 'none') {
    tasksContainer.style.display = 'block';
    if (chevron) chevron.style.transform = 'rotate(180deg)';
    groupEl.classList.add('expanded');
  }
  // Select the first task item inside the group
  const firstItem = tasksContainer?.querySelector('.todo-item');
  if (firstItem) {
    clearKeyboardSelection();
    keyboardSelectedTaskId = firstItem.dataset.todoId;
    if (multiSelect && !selectedTodoIds.has(keyboardSelectedTaskId)) {
      isMultiSelectMode = true;
      toggleTodoSelection(keyboardSelectedTaskId);
    }
    highlightKeyboardSelectedTask();
    firstItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
}

function extendSelectionUp() {
  if (!keyboardSelectedTaskId || !currentKeyboardColumn) return;

  const columnId = currentKeyboardColumn === 'action' ? 'actionColumn' : 'fyiColumn';
  const columnEl = document.getElementById(columnId);
  if (!columnEl) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedTaskId);
  }

  // Find the current task element
  const currentTask = columnEl.querySelector(`.todo-item[data-todo-id="${keyboardSelectedTaskId}"]`);
  if (!currentTask) return;

  // Find the list container
  const listContainer = currentTask.closest('.fyi-list, .todos-list') || currentTask.parentElement;
  if (!listContainer) return;

  // Determine the "top-level" element for the current task
  const currentTopLevel = currentTask.closest('.task-group') || currentTask;

  // If current task is inside an expanded task-group, check if there's a previous sibling within the group
  const taskGroupTasks = currentTask.closest('.task-group-tasks');
  if (taskGroupTasks) {
    const siblingsInGroup = Array.from(taskGroupTasks.querySelectorAll('.todo-item'));
    const idxInGroup = siblingsInGroup.indexOf(currentTask);
    if (idxInGroup > 0) {
      // Move to previous item within the same group
      const prevInGroup = siblingsInGroup[idxInGroup - 1];
      keyboardSelectedTaskId = prevInGroup.dataset.todoId;
      if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
        toggleTodoSelection(keyboardSelectedTaskId);
      }
      highlightKeyboardSelectedTask();
      prevInGroup.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
      return;
    }
    // At first item in group — fall through to look at previous top-level sibling
  }

  // Get all top-level children of the list
  const topLevelItems = Array.from(listContainer.children).filter(el =>
    (el.classList.contains('todo-item') || el.classList.contains('task-group')) &&
    !el.closest('.meetings-accordion') &&
    !el.closest('.documents-accordion')
  );

  const topIndex = topLevelItems.indexOf(currentTopLevel);
  if (topIndex <= 0) return;

  // Look at the previous top-level item
  const prevTopLevel = topLevelItems[topIndex - 1];

  if (prevTopLevel.classList.contains('task-group')) {
    // It's a task group — expand if needed, select LAST item
    const tasksContainer = prevTopLevel.querySelector('.task-group-tasks');
    const chevron = prevTopLevel.querySelector('.task-group-chevron');
    if (tasksContainer && tasksContainer.style.display === 'none') {
      tasksContainer.style.display = 'block';
      if (chevron) chevron.style.transform = 'rotate(180deg)';
      prevTopLevel.classList.add('expanded');
    }
    const allItems = tasksContainer?.querySelectorAll('.todo-item');
    const lastItem = allItems?.[allItems.length - 1];
    if (lastItem) {
      keyboardSelectedTaskId = lastItem.dataset.todoId;
      if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
        toggleTodoSelection(keyboardSelectedTaskId);
      }
      highlightKeyboardSelectedTask();
      lastItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    }
  } else if (prevTopLevel.classList.contains('todo-item')) {
    keyboardSelectedTaskId = prevTopLevel.dataset.todoId;
    if (!selectedTodoIds.has(keyboardSelectedTaskId)) {
      toggleTodoSelection(keyboardSelectedTaskId);
    }
    highlightKeyboardSelectedTask();
    prevTopLevel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }
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
  keyboardSelectedFeedId = null;
  document.querySelectorAll('.todo-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
  document.querySelectorAll('.note-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
  document.querySelectorAll('.dailyfeed-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
}

// Note selection functions
async function selectFirstFeedItem() {
  clearKeyboardSelection();
  clearMultiSelection();

  currentKeyboardColumn = 'feed';

  // Make sure Feed tab is active in col3
  const feedTab = document.getElementById('dailyFeedTab');
  if (feedTab && !feedTab.classList.contains('active')) {
    feedTab.click();
    // Wait for feed items to load (poll up to 3 seconds)
    let attempts = 0;
    while (attempts < 30) {
      await new Promise(r => setTimeout(r, 100));
      const fl = document.querySelector('.dailyfeed-list');
      if (fl && fl.querySelector('.dailyfeed-item')) break;
      attempts++;
    }
  }

  const feedList = document.querySelector('.dailyfeed-list');
  if (!feedList) return;

  const firstItem = feedList.querySelector('.dailyfeed-item');
  if (!firstItem) return;

  keyboardSelectedFeedId = firstItem.dataset.taskId;
  highlightKeyboardSelectedFeed();
  firstItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function selectFirstAllTasksItem() {
  clearKeyboardSelection();
  clearMultiSelection();

  currentKeyboardColumn = 'alltasks';

  // Make sure All Tasks tab is active in col3
  const allTasksTab = document.getElementById('allTasksTab');
  if (allTasksTab && !allTasksTab.classList.contains('active')) {
    allTasksTab.click();
    let attempts = 0;
    while (attempts < 30) {
      await new Promise(r => setTimeout(r, 100));
      const list = document.querySelector('.alltasks-list');
      if (list && list.querySelector('.todo-item')) break;
      attempts++;
    }
  }

  const list = document.querySelector('.alltasks-list');
  if (!list) return;

  const firstItem = list.querySelector('.todo-item');
  if (!firstItem) return;

  keyboardSelectedTaskId = firstItem.dataset.todoId;
  highlightKeyboardSelection();
  firstItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function highlightKeyboardSelectedFeed() {
  document.querySelectorAll('.dailyfeed-item.keyboard-selected').forEach(el => {
    el.classList.remove('keyboard-selected');
  });
  if (!keyboardSelectedFeedId) return;
  const item = document.querySelector(`.dailyfeed-item[data-task-id="${keyboardSelectedFeedId}"]`);
  if (item) {
    item.classList.add('keyboard-selected');
  }
}

function moveFeedSelection(direction) {
  if (!keyboardSelectedFeedId || currentKeyboardColumn !== 'feed') return;

  const feedList = document.querySelector('.dailyfeed-list');
  if (!feedList) return;

  const items = Array.from(feedList.querySelectorAll('.dailyfeed-item'));
  const currentIndex = items.findIndex(i => i.dataset.taskId === keyboardSelectedFeedId);
  if (currentIndex === -1) return;

  let newIndex;
  if (direction === 'down') {
    newIndex = currentIndex < items.length - 1 ? currentIndex + 1 : currentIndex;
  } else {
    newIndex = currentIndex > 0 ? currentIndex - 1 : 0;
  }

  keyboardSelectedFeedId = items[newIndex].dataset.taskId;
  highlightKeyboardSelectedFeed();
  items[newIndex].scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function extendFeedSelectionDown() {
  if (!keyboardSelectedFeedId || currentKeyboardColumn !== 'feed') return;

  const feedList = document.querySelector('.dailyfeed-list');
  if (!feedList) return;

  const items = Array.from(feedList.querySelectorAll('.dailyfeed-item'));
  const currentIndex = items.findIndex(i => i.dataset.taskId === keyboardSelectedFeedId);
  if (currentIndex === -1 || currentIndex >= items.length - 1) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedFeedId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedFeedId);
  }

  // Move down
  const nextItem = items[currentIndex + 1];
  keyboardSelectedFeedId = nextItem.dataset.taskId;
  highlightKeyboardSelectedFeed();
  nextItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  // Add new item to selection
  if (!selectedTodoIds.has(keyboardSelectedFeedId)) {
    toggleTodoSelection(keyboardSelectedFeedId);
  }

  // Update feed item visual state
  updateFeedSelectionUI();
}

function extendFeedSelectionUp() {
  if (!keyboardSelectedFeedId || currentKeyboardColumn !== 'feed') return;

  const feedList = document.querySelector('.dailyfeed-list');
  if (!feedList) return;

  const items = Array.from(feedList.querySelectorAll('.dailyfeed-item'));
  const currentIndex = items.findIndex(i => i.dataset.taskId === keyboardSelectedFeedId);
  if (currentIndex === -1 || currentIndex <= 0) return;

  // Add current to multi-selection if not already
  if (!selectedTodoIds.has(keyboardSelectedFeedId)) {
    isMultiSelectMode = true;
    toggleTodoSelection(keyboardSelectedFeedId);
  }

  // Move up
  const prevItem = items[currentIndex - 1];
  keyboardSelectedFeedId = prevItem.dataset.taskId;
  highlightKeyboardSelectedFeed();
  prevItem.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

  // Add new item to selection
  if (!selectedTodoIds.has(keyboardSelectedFeedId)) {
    toggleTodoSelection(keyboardSelectedFeedId);
  }

  // Update feed item visual state
  updateFeedSelectionUI();
}

function updateFeedSelectionUI() {
  document.querySelectorAll('.dailyfeed-item').forEach(item => {
    const taskId = String(item.dataset.taskId);
    if (selectedTodoIds.has(taskId)) {
      item.classList.add('multi-selected');
    } else {
      item.classList.remove('multi-selected');
    }
  });
}

async function markFeedItemDone(feedId) {
  const taskId = parseInt(feedId);
  const item = document.querySelector(`.dailyfeed-item[data-task-id="${feedId}"]`);
  if (item) {
    item.classList.add('completing');
    setTimeout(() => {
      item.remove();
      // Check if date group is now empty
      document.querySelectorAll('.dailyfeed-date-group').forEach(group => {
        if (group.querySelectorAll('.dailyfeed-item').length === 0) group.remove();
      });
      if (typeof window.updateDailyFeedCount === 'function') window.updateDailyFeedCount();
      const feedList = document.querySelector('.dailyfeed-list');
      if (feedList && feedList.querySelectorAll('.dailyfeed-item').length === 0) {
        feedList.style.display = 'none';
        const emptyState = document.getElementById('dailyFeedEmpty');
        if (emptyState) emptyState.style.display = 'flex';
      }
    }, 400);
  }
  allTodos = allTodos.filter(t => t.id != taskId);
  keyboardSelectedFeedId = null;
  showToastNotification('Feed item marked as done');
  updateTodoField(taskId, 'status', 1).catch(err => console.error('Error marking feed item done:', err));
}

async function selectFirstNoteInColumn() {
  clearKeyboardSelection();
  clearMultiSelection();

  currentKeyboardColumn = 'notes';

  // Make sure Notes tab is active
  const notesTab = document.getElementById('notesTab');
  if (notesTab && !notesTab.classList.contains('active')) {
    notesTab.click();
    // Wait for notes to load (poll for note items up to 3 seconds)
    let attempts = 0;
    while (attempts < 30) {
      await new Promise(r => setTimeout(r, 100));
      const notesContent = document.getElementById('notesContent');
      if (notesContent && notesContent.querySelector('.note-item')) break;
      attempts++;
    }
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
    const parentGroup = task.closest('.task-group');
    task.remove();
    // Clean up empty task group
    if (parentGroup) {
      const remaining = parentGroup.querySelectorAll('.todo-item:not(.completing)');
      if (remaining.length === 0) {
        parentGroup.remove();
      } else {
        const countEl = parentGroup.querySelector('.task-group-count');
        if (countEl) {
          countEl.textContent = `${remaining.length} item${remaining.length > 1 ? 's' : ''}`;
        }
      }
    }
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

// ============================================
// Rich Text Formatting for contenteditable reply inputs
// ============================================
function setupRichTextFormatting(input) {
  if (!document.getElementById('richtext-styles')) {
    const style = document.createElement('style');
    style.id = 'richtext-styles';
    style.textContent = `
      .transcript-reply-input ol, .transcript-reply-input ul {
        margin: 4px 0; padding-left: 24px;
      }
      .transcript-reply-input li { margin: 2px 0; }
      .transcript-reply-input b, .transcript-reply-input strong { font-weight: 700; }
      .transcript-reply-input i, .transcript-reply-input em { font-style: italic; }
      .transcript-reply-input a, .feed-oracle-reply a { color: #667eea; text-decoration: underline; cursor: pointer; }
      body.dark-mode .transcript-reply-input a, body.dark-mode .feed-oracle-reply a { color: #8fa4f8; }
    `;
    document.head.appendChild(style);
  }

  // Bold / Italic shortcuts
  input.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'b' || e.key === 'B')) {
      e.preventDefault();
      document.execCommand('bold', false, null);
      return;
    }
    if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'i' || e.key === 'I')) {
      e.preventDefault();
      document.execCommand('italic', false, null);
      return;
    }

    // Enter in list: continue list or exit on empty item
    if (e.key === 'Enter' && !e.metaKey && !e.ctrlKey && !e.shiftKey) {
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      const node = range.startContainer;
      const li = node.nodeType === Node.TEXT_NODE
        ? node.parentElement.closest('li')
        : node.closest?.('li');
      if (!li) return;

      const list = li.closest('ol, ul');
      if (!list) return;

      if (li.textContent.trim() === '') {
        // Empty bullet — exit the list
        e.preventDefault();
        li.remove();
        const div = document.createElement('div');
        div.innerHTML = '<br>';
        list.parentNode.insertBefore(div, list.nextSibling);
        if (list.querySelectorAll('li').length === 0) {
          list.remove();
        }
        const newRange = document.createRange();
        newRange.setStart(div, 0);
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);
      } else {
        // Non-empty bullet — create next list item
        e.preventDefault();

        // Split text at cursor if needed
        const newLi = document.createElement('li');
        
        // If cursor is at the end of the li, just create empty new li
        // If cursor is in the middle, move the rest to the new li
        if (range.endOffset < (node.nodeType === Node.TEXT_NODE ? node.textContent.length : node.childNodes.length)) {
          // Split: move content after cursor to new li
          const afterRange = document.createRange();
          afterRange.setStart(range.endContainer, range.endOffset);
          afterRange.setEndAfter(li.lastChild || li);
          const fragment = afterRange.extractContents();
          newLi.appendChild(fragment);
        }
        
        // If new li is empty, add a br so cursor can be placed
        if (!newLi.hasChildNodes() || newLi.textContent === '') {
          newLi.innerHTML = '<br>';
        }

        // Insert after current li
        li.parentNode.insertBefore(newLi, li.nextSibling);

        // Place cursor at the start of the new li
        const newRange = document.createRange();
        newRange.setStart(newLi, 0);
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);
      }
    }
  });

  // Detect "1. " or "- " patterns AFTER typing (on input event)
  input.addEventListener('input', (e) => {
    if (e.inputType !== 'insertText' || e.data !== ' ') return;

    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    let node = range.startContainer;
    if (node.nodeType !== Node.TEXT_NODE) return;
    if (node.parentElement.closest('ol, ul')) return;

    const fullText = node.textContent;
    const offset = range.startOffset;
    // Get text from start of this text node up to cursor
    const beforeCursor = fullText.substring(0, offset);

    // Check patterns: "1. " or "- " (the space was just typed)
    const olMatch = beforeCursor.match(/^(\d+)\.\s$/);
    const ulMatch = beforeCursor.match(/^-\s$/);

    if (!olMatch && !ulMatch) return;

    // Remove trigger text and trailing content stays
    const afterCursor = fullText.substring(offset);

    // Use a timeout to let the browser finish processing the input event
    setTimeout(() => {
      // Select all content in this text node and delete it
      node.textContent = afterCursor || '\u200B'; // Zero-width space if empty

      // Position cursor at start
      const newRange = document.createRange();
      newRange.setStart(node, 0);
      newRange.collapse(true);
      sel.removeAllRanges();
      sel.addRange(newRange);

      // Now apply the list
      if (olMatch) {
        document.execCommand('insertOrderedList', false, null);
      } else {
        document.execCommand('insertUnorderedList', false, null);
      }

      // Clean up zero-width space if it's still there
      const activeLi = sel.anchorNode?.nodeType === Node.TEXT_NODE ? sel.anchorNode : sel.anchorNode?.firstChild;
      if (activeLi && activeLi.textContent === '\u200B') {
        activeLi.textContent = '';
      }
    }, 0);
  });

  // Cmd/Ctrl+K: Insert hyperlink
  input.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'k' || e.key === 'K')) {
      e.preventDefault();
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      const selectedText = range.toString();
      if (!selectedText) {
        const url = prompt('Enter URL:');
        if (!url) return;
        const text = prompt('Link text:', url) || url;
        const a = document.createElement('a');
        a.href = url; a.target = '_blank'; a.rel = 'noopener noreferrer';
        a.textContent = text;
        range.insertNode(a);
        const space = document.createTextNode('\u00A0');
        a.parentNode.insertBefore(space, a.nextSibling);
        const nr = document.createRange(); nr.setStartAfter(space); nr.collapse(true);
        sel.removeAllRanges(); sel.addRange(nr);
      } else {
        const url = prompt('Enter URL for "' + selectedText + '":');
        if (!url) return;
        document.execCommand('createLink', false, url);
        const newLink = sel.anchorNode?.parentElement?.closest('a') || input.querySelector(`a[href="${url}"]`);
        if (newLink) { newLink.target = '_blank'; newLink.rel = 'noopener noreferrer'; }
        sel.collapseToEnd();
      }
    }
  });

  // Paste URL on selected text → auto-hyperlink
  if (window.OracleNewMessage?.setupPasteToHyperlink) {
    window.OracleNewMessage.setupPasteToHyperlink(input);
  }
  if (window.OracleNewMessage?.setupEmoticonReplace) {
    window.OracleNewMessage.setupEmoticonReplace(input);
  }
}

// Convert contenteditable HTML to Slack mrkdwn format
function convertContentEditableToSlackMrkdwn(input) {
  const processNode = (node) => {
    if (node.nodeType === Node.TEXT_NODE) {
      return node.textContent;
    }
    if (node.nodeName === 'BR') return '\n';

    // Mention tags
    if (node.classList?.contains('mention-tag')) {
      const slackId = node.dataset.slackId;
      return slackId ? `<@${slackId}>` : node.textContent;
    }

    // Links → Slack format: <url|text>
    if (node.nodeName === 'A') {
      const href = node.getAttribute('href') || '';
      let inner = '';
      node.childNodes.forEach(child => { inner += processNode(child); });
      if (inner.trim() === href.trim()) return `<${href}>`;
      return `<${href}|${inner}>`;
    }

    // Process children first
    let inner = '';
    node.childNodes.forEach(child => { inner += processNode(child); });

    const tag = node.nodeName.toUpperCase();

    // Bold
    if (tag === 'B' || tag === 'STRONG') return `*${inner}*`;
    // Italic
    if (tag === 'I' || tag === 'EM') return `_${inner}_`;
    // Strikethrough
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') return `~${inner}~`;

    // List items
    if (tag === 'LI') {
      const parentList = node.closest('ol, ul');
      if (parentList?.nodeName === 'OL') {
        const items = Array.from(parentList.querySelectorAll(':scope > li'));
        const idx = items.indexOf(node) + 1;
        return `${idx}. ${inner.trim()}`;
      }
      return `• ${inner.trim()}`;
    }

    // Lists — join items with newlines
    if (tag === 'OL' || tag === 'UL') {
      const items = [];
      node.querySelectorAll(':scope > li').forEach(li => {
        items.push(processNode(li));
      });
      return '\n' + items.join('\n') + '\n';
    }

    // Div/P — treat as newline-separated blocks
    if (tag === 'DIV' || tag === 'P') {
      // Only add newline if there's content before this
      return '\n' + inner;
    }

    return inner;
  };

  let result = '';
  input.childNodes.forEach(node => { result += processNode(node); });
  // Clean up: collapse multiple newlines, trim
  return result.replace(/\n{3,}/g, '\n\n').trim();
}

// Clear selection when clicking outside
document.addEventListener('click', (e) => {
  // Don't clear selection if clicking on:
  // - A todo item
  // - The bulk action button
  // - The transcript slider overlay or its contents
  // - The attachment preview modal
  // - While holding Command/Ctrl/Shift key
  if (!e.target.closest('.todo-item') &&
    !e.target.closest('.meeting-item') &&
    !e.target.closest('.bulk-action-btn') &&
    !e.target.closest('.transcript-slider-overlay') &&
    !e.target.closest('.attachment-preview-modal') &&
    !e.target.closest('.due-by-menu') &&
    !isCommandKeyPressed &&
    !e.shiftKey) {
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

        // Move slider - 5 tabs now
        if (col3TabSlider) {
          col3TabSlider.style.width = `calc(20% - 2px)`;
          col3TabSlider.style.left = `calc(${index * 20}% + 3px)`;
        }

        // Switch content
        const targetId = tab.dataset.col3;
        document.querySelectorAll('.col3-content').forEach(content => {
          content.classList.remove('active');
        });

        if (targetId === 'dailyfeed') {
          const dailyFeedContent = document.getElementById('dailyFeedContent');
          dailyFeedContent?.classList.add('active');
          if (addNoteBtn) addNoteBtn.style.display = 'none';
          loadDailyFeed();
        } else if (targetId === 'notes') {
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
          // Hide the action button for All Tasks
          if (addNoteBtn) addNoteBtn.style.display = 'none';
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
      else if (targetId === 'dailyfeed') loadDailyFeed(true);
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
          // If is_latest_from_self is true, always treat as read
          if (task.is_latest_from_self === true) {
            readTaskIds.add(idStr);
            if (task.updated_at) previousTaskTimestamps.set(idStr, task.updated_at);
            return;
          }
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
      // But skip items where is_latest_from_self is true
      updatedItems.forEach(t => {
        if (t.is_latest_from_self === true) return;
        const idStr = String(t.id);
        readTaskIds.delete(idStr);
        console.log(`📝 Task ${t.id} marked as unread (updated_at changed)`);
      });

      // Force-read any is_latest_from_self items (including new ones)
      newTodos.forEach(t => {
        if (t.is_latest_from_self === true) {
          readTaskIds.add(String(t.id));
        }
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
          // If is_latest_from_self is true, always treat as read
          if (task.is_latest_from_self === true) {
            readTaskIds.add(idStr);
            if (task.updated_at) previousTaskTimestamps.set(idStr, task.updated_at);
            return;
          }
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
      // But skip items where is_latest_from_self is true
      updatedItems.forEach(t => {
        if (t.is_latest_from_self === true) return;
        const idStr = String(t.id);
        readTaskIds.delete(idStr);
        console.log(`📝 FYI Task ${t.id} marked as unread (updated_at changed)`);
      });

      // Force-read any is_latest_from_self items (including new ones)
      newFyiItems.forEach(t => {
        if (t.is_latest_from_self === true) {
          readTaskIds.add(String(t.id));
        }
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

    // Build Slack channel groups as flat items (same as Action section - no nested accordion)
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
    // Also remove any existing flat FYI slack groups
    container.querySelectorAll('.fyi-slack-group').forEach(el => el.remove());

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

    // Build unified sorted list of all FYI items (Slack groups + tag groups + individual tasks)
    const unifiedFyiItems = [];

    // Add Slack channel groups
    Object.values(fyiSlackGroups).forEach(group => {
      unifiedFyiItems.push({
        type: 'slack-group',
        data: group,
        sortTimestamp: new Date(group.latestUpdate)
      });
    });

    // Add tag groups
    Object.values(fyiTagGroups).forEach(group => {
      unifiedFyiItems.push({
        type: 'tag-group',
        data: group,
        sortTimestamp: new Date(group.latestUpdate)
      });
    });

    // Add remaining individual tasks
    itemsForIndividualDisplay.forEach(task => {
      unifiedFyiItems.push({
        type: 'task',
        data: task,
        sortTimestamp: new Date(task.updated_at || task.created_at)
      });
    });

    // Sort: unread items first, then by timestamp descending (same as Action section)
    unifiedFyiItems.sort((a, b) => {
      const aHasUnread = a.type === 'task'
        ? isTaskUnread(a.data.id)
        : a.data.tasks.some(t => isTaskUnread(t.id));
      const bHasUnread = b.type === 'task'
        ? isTaskUnread(b.data.id)
        : b.data.tasks.some(t => isTaskUnread(t.id));
      if (aHasUnread && !bHasUnread) return -1;
      if (!aHasUnread && bHasUnread) return 1;
      return b.sortTimestamp - a.sortTimestamp;
    });

    // Build HTML in sorted order
    fyiList.innerHTML = unifiedFyiItems.map((item, index) => {
      if (item.type === 'slack-group') {
        return buildSlackChannelGroupHtml(item.data, newItemIds);
      } else if (item.type === 'tag-group') {
        return buildFyiTagGroupHtml(item.data, newItemIds);
      } else {
      const task = item.data;
      const messageLink = task.message_link || '';
      const secondaryLinks = task.secondary_links || [];
      const isNewItem = newItemIds.includes(task.id);

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

      const dueByHtml = task.due_by ? '<span class="todo-due ' + (new Date(task.due_by) < new Date() ? 'overdue' : '') + '">' + formatDueBy(task.due_by) + '</span>' : '';

      // Handle descriptions - always show 2 lines with View more
      const taskNameRaw = task.task_name || '';
      const taskNameEscaped = escapeHtml(taskNameRaw);
      const maxLength = 180;
      const hasTitle = !!task.task_title;
      const needsViewMore = taskNameRaw.length > maxLength || (hasTitle && taskNameRaw.length > 60);
      let taskNameHtml;

      if (needsViewMore) {
        const truncated = escapeHtml(taskNameRaw.substring(0, maxLength));
        taskNameHtml = '<span class="todo-text-content" data-full-text="' + taskNameEscaped + '">' + taskNameEscaped + '</span><span class="view-more-inline" data-todo-id="' + task.id + '">View more</span>';
      } else {
        taskNameHtml = '<span class="todo-text-content">' + taskNameEscaped + '</span>';
      }

      // Render tags for each item (filter out invalid tags like "null")
      const itemTags = (task.tags || []).filter(isValidTag);
      const tagsHtml = itemTags.length > 0 ? `
          <div class="todo-tags" style="display: flex; flex-wrap: wrap; gap: 4px; margin-top: 6px;">
            ${itemTags.map(tag => '<span class="todo-tag" data-tag="' + escapeHtml(tag) + '" style="display: inline-block; background: rgba(102, 126, 234, 0.1); color: #667eea; padding: 2px 6px; border-radius: 4px; font-size: 10px; font-weight: 500; cursor: pointer; border: 1px solid rgba(102, 126, 234, 0.3); transition: all 0.2s;">' + escapeHtml(tag) + '</span>').join('')}
          </div>
        ` : '';

      // Participant text (e.g., channel name)
      const participantHtml = task.participant_text ? '<div class="todo-participant">' + escapeHtml(task.participant_text) + '</div>' : '';
      const typeIconHtml = getTypeIconHtml(task.type || null, task.participant_text || '');

      // Check if task is unread
      const isUnread = isTaskUnread(task.id, task.updated_at);

      // Eye icon for Slack FYI tasks (ungrouped)
      const fyiSlackChannelId = extractSlackChannelId(messageLink);
      const fyiChannelName = task.participant_text || '';
      const fyiEyeIconHtml = fyiSlackChannelId
        ? `<span class="todo-slack-eye-btn" data-channel-id="${escapeHtml(fyiSlackChannelId)}" data-channel-name="${escapeHtml(fyiChannelName)}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;">#</span>`
        : '';

      return `<div class="todo-item fyi-item ${isNewItem ? 'appearing' : ''} ${isUnread ? 'unread' : ''}" style="animation-delay: ${index * 0.05}s" data-todo-id="${task.id}">
          ${getTrendingIconHtml(task)}${fyiEyeIconHtml}
          <div class="todo-left-actions">
            <div class="todo-checkbox" data-todo-id="${task.id}"></div>
          </div>
          <div class="todo-content">
            ${task.task_title ? '<div class="todo-title">' + escapeHtml(task.task_title) + '</div>' : ''}
            <div class="todo-text${needsViewMore ? ' truncated' : ''}">${taskNameHtml}</div>
            ${participantHtml}
            <div class="todo-meta">
              <span class="todo-date">${formatDate(task.updated_at || task.created_at)}</span>
              ${dueByHtml}
              ${messageLink ? '<a href="' + messageLink + '" target="_blank" class="todo-source" title="' + sourceTitle + '" style="padding: 6px;">' + sourceIconHtml + '</a>' : ''}
              ${secondaryLinksHtml}
              ${(typeIconHtml || fyiSlackChannelId) ? `<span style="flex:1;min-width:4px;"></span>` : ''}
              ${typeIconHtml || ''}
              ${fyiSlackChannelId ? `<span class="slack-bell-channel-btn" data-channel-id="${escapeHtml(fyiSlackChannelId)}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-left:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
            </div>
            ${tagsHtml}
          </div>
          ${messageLink && messageLink.includes('mail.google.com') ? `<span class="todo-block-email-btn" data-todo-id="${task.id}" title="Block email participants"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
        </div>`;
      }
    }).join('');

    // Add event listeners - make whole item clickable for transcript or multi-select
    fyiList.querySelectorAll('.todo-item').forEach(item => {
      const todoId = item.dataset.todoId;

      // Block email btn for FYI email items
      const fyiBlockBtn = item.querySelector('.todo-block-email-btn');
      if (fyiBlockBtn) {
        fyiBlockBtn.addEventListener('click', (e) => { e.stopPropagation(); showBlockEmailModal(fyiBlockBtn.dataset.todoId); });
      }

      // Block Slack channel btn for FYI items
      const fyiBlockChBtn = item.querySelector('.slack-bell-channel-btn');
      if (fyiBlockChBtn && !fyiBlockChBtn.dataset.blockHandlerAttached) {
        fyiBlockChBtn.dataset.blockHandlerAttached = 'true';
        fyiBlockChBtn.addEventListener('click', (e) => { e.stopPropagation(); showBellDropdown(e, fyiBlockChBtn.dataset.channelId); });
      }

      // Eye btn for FYI Slack items
      const fyiEyeBtn = item.querySelector('.todo-slack-eye-btn');
      fyiEyeBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        const channelId = fyiEyeBtn.dataset.channelId;
        const chName = fyiEyeBtn.dataset.channelName;
        if (channelId) openChannelTranscriptSlider(channelId, chName);
      });

      item.addEventListener('click', (e) => {
        if (e.target.closest('.todo-checkbox') ||
          e.target.closest('.todo-clock') ||
          e.target.closest('.todo-slack-eye-btn') ||
          e.target.closest('.todo-block-email-btn') ||
          e.target.closest('.slack-bell-channel-btn') ||
          e.target.closest('.todo-source') ||
          e.target.closest('.todo-tag') ||
          e.target.closest('.view-more-inline') ||
          e.target.closest('a')) {
          return;
        }

        // Command/Ctrl/Shift + click on the item itself triggers multi-select
        if (e.metaKey || e.ctrlKey || e.shiftKey) {
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
        // Check if Command/Ctrl/Shift key is pressed for multi-select
        if (e.metaKey || e.ctrlKey || e.shiftKey) {
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

    // Setup event handlers for Slack channel groups and tag groups now inside fyiList
    setupSlackChannelGroups(fyiList);
    setupFyiTagGroups(fyiList);
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

      const _fyiEyeId = extractSlackChannelId(taskMessageLink);
      const _fyiEyeHtml = _fyiEyeId ? `<span class="todo-slack-eye-btn" data-channel-id="${escapeHtml(_fyiEyeId)}" data-channel-name="${escapeHtml(task.participant_text || '')}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;">#</span>` : '';
      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                ${getTrendingIconHtml(task)}${_fyiEyeHtml}
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
                    <div class="todo-sources" style="display: flex; align-items: center; gap: 2px; flex: 1; min-width: 0;">
                      ${sourceIconHtml}
                      ${secondaryLinksHtml}
                      ${(getTypeIconHtml(task.type || null, task.participant_text || '') || _fyiEyeId) ? `<span style="flex:1;min-width:4px;display:inline-flex;"></span>` : ''}
                      ${getTypeIconHtml(task.type || null, task.participant_text || '') || ''}
                      ${_fyiEyeId ? `<span class="slack-bell-channel-btn" data-channel-id="${escapeHtml(_fyiEyeId)}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-left:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
                    </div>
                  </div>
                  ${tagsHtml}
                </div>
                ${taskMessageLink && taskMessageLink.includes('mail.google.com') ? `<span class="todo-block-email-btn" data-todo-id="${task.id}" title="Block email participants"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
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
        // Eye btn — open channel transcript slider
        const eyeBtn = taskItem.querySelector('.todo-slack-eye-btn');
        if (eyeBtn) {
          eyeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const channelId = eyeBtn.dataset.channelId;
            const chName = eyeBtn.dataset.channelName;
            if (channelId) openChannelTranscriptSlider(channelId, chName);
          });
        }

        // Block email btn — show block modal
        const blockBtn = taskItem.querySelector('.todo-block-email-btn');
        if (blockBtn) {
          blockBtn.addEventListener('click', (e) => { e.stopPropagation(); showBlockEmailModal(blockBtn.dataset.todoId); });
        }

        // Block Slack channel btn
        const blockChBtnAction = taskItem.querySelector('.slack-bell-channel-btn');
        if (blockChBtnAction && !blockChBtnAction.dataset.blockHandlerAttached) {
          blockChBtnAction.dataset.blockHandlerAttached = 'true';
          blockChBtnAction.addEventListener('click', (e) => { e.stopPropagation(); showBellDropdown(e, blockChBtnAction.dataset.channelId); });
        }

        taskItem.addEventListener('click', (e) => {
          // Don't trigger if clicking on checkbox, view-more, eye btn, block btn, or links
          if (e.target.closest('.todo-checkbox') ||
            e.target.closest('.view-more-inline') ||
            e.target.closest('.todo-slack-eye-btn') ||
            e.target.closest('.todo-block-email-btn') ||
            e.target.closest('.slack-bell-channel-btn') ||
            e.target.closest('a')) {
            return;
          }

          const todoId = taskItem.dataset.todoId;
          if (!todoId) return;

          if (e.metaKey || e.ctrlKey || e.shiftKey) {
            e.preventDefault();
            isMultiSelectMode = true;
            toggleTodoSelection(todoId);
            return;
          }
          showTranscriptSlider(todoId);
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
    const restoreSvg = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>';
    tasksList.querySelectorAll('.todo-checkbox').forEach(cb => { cb.style.display = 'none'; const la = cb.closest('.todo-left-actions'); if (la) la.style.display = 'none'; const item = cb.closest('.todo-item'); if (!item) return; const b = document.createElement('button'); b.className = 'alltask-reactivate'; b.dataset.taskId = cb.dataset.todoId; b.title = 'Mark as active'; b.innerHTML = restoreSvg; b.style.cssText = 'position:absolute;bottom:8px;right:8px;'; item.appendChild(b); });
    tasksList.querySelectorAll('.meeting-checkbox').forEach(cb => { cb.style.display = 'none'; const item = cb.closest('.todo-item,.meeting-item'); if (!item) return; item.style.position = 'relative'; const b = document.createElement('button'); b.className = 'alltask-reactivate'; b.dataset.taskId = cb.dataset.todoId; b.title = 'Mark as active'; b.innerHTML = restoreSvg; b.style.cssText = 'position:absolute;bottom:8px;right:8px;'; item.appendChild(b); });
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
  // ============================================
  // DAILY FEED
  // ============================================
  const DAILY_FEED_TAG_ID = '2282';
  const DAILY_FEED_TAG_NAME = 'Daily Feed';

  async function loadDailyFeed(forceRefresh = false) {
    const container = document.querySelector('.dailyfeed-container');
    const emptyState = container?.querySelector('.empty-state');
    const feedList = container?.querySelector('.dailyfeed-list');
    if (emptyState) emptyState.style.display = 'none';
    if (feedList) feedList.style.display = 'none';
    showLoader(container);

    try {
      // If allTodos not loaded yet or force refresh, fetch them
      if (!allTodos || allTodos.length === 0 || forceRefresh) {
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
        const responseText = await response.text();
        if (responseText && responseText.trim()) {
          const data = JSON.parse(responseText);
          allTodos = Array.isArray(data) ? data : (data?.todos || data?.results || []);
        }
      }

      // Filter tasks with Daily Feed tag (match both ID and name)
      const feedItems = (allTodos || []).filter(task => {
        const tags = task.tags || [];
        return tags.some(t => 
          String(t) === DAILY_FEED_TAG_ID || 
          String(t).toLowerCase() === DAILY_FEED_TAG_NAME.toLowerCase()
        );
      }).sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

      hideLoader(container);
      if (feedItems.length > 0) {
        displayDailyFeed(feedItems);
      } else {
        if (feedList) feedList.style.display = 'none';
        if (emptyState) emptyState.style.display = 'flex';
        const feedCountEl = document.getElementById('dailyFeedCount');
        if (feedCountEl) feedCountEl.style.display = 'none';
      }
    } catch (error) {
      console.error('Error loading daily feed:', error);
      hideLoader(container);
      if (emptyState) {
        emptyState.style.display = 'flex';
        emptyState.innerHTML = '<h3>Error: ' + error.message + '</h3>';
      }
    }
  }

  function updateDailyFeedCount() {
    const feedCountEl = document.getElementById('dailyFeedCount');
    if (!feedCountEl) return;
    const feedList = document.querySelector('.dailyfeed-list');
    const count = feedList ? feedList.querySelectorAll('.dailyfeed-item:not(.completing)').length : 0;
    if (count > 0) {
      feedCountEl.textContent = count;
      feedCountEl.style.display = 'inline';
    } else {
      feedCountEl.style.display = 'none';
    }
  }

  function displayDailyFeed(feedItems) {
    const container = document.querySelector('.dailyfeed-container');
    let feedList = container?.querySelector('.dailyfeed-list');
    const emptyState = container?.querySelector('.empty-state');
    if (emptyState) emptyState.style.display = 'none';

    if (!feedList) {
      feedList = document.createElement('div');
      feedList.className = 'dailyfeed-list';
      container.appendChild(feedList);
    }

    // Group by date
    const groupedByDate = {};
    feedItems.forEach(item => {
      const date = new Date(item.created_at).toLocaleDateString('en-US', { 
        weekday: 'short', month: 'short', day: 'numeric' 
      });
      if (!groupedByDate[date]) groupedByDate[date] = [];
      groupedByDate[date].push(item);
    });

    feedList.innerHTML = Object.entries(groupedByDate).map(([date, items]) => `
      <div class="dailyfeed-date-group">
        <div style="font-size: 11px; font-weight: 600; color: var(--text-muted); text-transform: uppercase; letter-spacing: 0.5px; padding: 8px 4px 4px; border-bottom: 1px solid var(--border-light); margin-bottom: 8px;">${date}</div>
        ${items.map(item => {
          const domainRaw = item.message_link ? (() => { try { return new URL(item.message_link).hostname.replace('www.', ''); } catch(e) { return ''; } })() : '';
          const isSlackLink = domainRaw.includes('slack.com');
          const slackIconFeed = isSlackLink ? chrome.runtime.getURL('icon-slack.png') : '';
          const domain = domainRaw;
          const timeAgo = formatTimeAgoFresh(item.created_at);
          const descText = item.task_name || '';
          const needsTruncate = descText.length > 120;
          // If is_latest_from_self is true, always treat as read
          const isSelfLatest = item.is_latest_from_self === true;
          const isUnread = isSelfLatest ? false : isTaskUnread(item.id);
          // If is_latest_from_self, ensure it's marked as read in state
          if (isSelfLatest && !readTaskIds.has(String(item.id))) {
            readTaskIds.add(String(item.id));
          }
          return `
            <div class="dailyfeed-item ${isUnread ? 'unread' : ''}" data-task-id="${item.id}" data-status="${item.status}">
              <div class="dailyfeed-item-row">
                <div class="dailyfeed-checkbox" data-task-id="${item.id}" title="Mark as done"></div>
                <div class="dailyfeed-item-content">
                  <div class="dailyfeed-item-title">${escapeHtml(item.task_title || '')}</div>
                  ${descText ? `<div class="dailyfeed-item-desc ${needsTruncate ? 'truncated' : ''}">${escapeHtml(descText)}${needsTruncate ? `<span class="dailyfeed-view-more" data-task-id="${item.id}">View more</span>` : ''}</div>` : ''}
                  <div class="dailyfeed-item-meta">
                    ${item.message_link ? `<a class="dailyfeed-item-source" href="${escapeHtml(item.message_link)}" target="_blank" title="${escapeHtml(item.message_link)}">${isSlackLink ? `<img src="${slackIconFeed}" alt="Slack" style="width:14px;height:14px;object-fit:contain;vertical-align:middle;">` : escapeHtml(domain)}</a>` : ''}
                    <span class="dailyfeed-item-time">${timeAgo}</span>
                    <button class="dailyfeed-oracle-btn" data-task-id="${item.id}" title="Ask Oracle Assistant">∞</button>
                  </div>
                </div>
              </div>
            </div>
          `;
        }).join('')}
      </div>
    `).join('');

    feedList.style.display = 'flex';
    updateDailyFeedCount();

    // Checkbox click - mark as done
    feedList.querySelectorAll('.dailyfeed-checkbox').forEach(checkbox => {
      checkbox.addEventListener('click', async (e) => {
        e.stopPropagation();
        const taskId = parseInt(checkbox.dataset.taskId);
        checkbox.classList.add('checked');
        checkbox.innerHTML = '✓';
        const item = checkbox.closest('.dailyfeed-item');
        if (item) {
          item.classList.add('completing');
          // Remove from DOM after animation
          setTimeout(() => {
            item.remove();
            // Check if date group is now empty and remove it
            feedList.querySelectorAll('.dailyfeed-date-group').forEach(group => {
              if (group.querySelectorAll('.dailyfeed-item').length === 0) {
                group.remove();
              }
            });
            updateDailyFeedCount();
            // Show empty state if no items left
            if (feedList.querySelectorAll('.dailyfeed-item').length === 0) {
              feedList.style.display = 'none';
              const emptyState = document.getElementById('dailyFeedEmpty');
              if (emptyState) emptyState.style.display = 'flex';
            }
          }, 400);
        }
        // Update local arrays
        allTodos = allTodos.filter(t => t.id != taskId);
        showToastNotification('Feed item marked as done');
        // Send to backend in background
        updateTodoField(taskId, 'status', 1).catch(err => {
          console.error('Error marking feed item done:', err);
        });
      });
    });

    // View more / View less toggle
    feedList.querySelectorAll('.dailyfeed-view-more').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const desc = btn.closest('.dailyfeed-item-desc');
        if (!desc) return;
        if (desc.classList.contains('truncated')) {
          desc.classList.remove('truncated');
          btn.textContent = 'View less';
          btn.classList.add('view-less');
        } else {
          desc.classList.add('truncated');
          btn.textContent = 'View more';
          btn.classList.remove('view-less');
        }
      });
    });

    // Click on item title to open source
    // Oracle assistant button click - open dedicated feed oracle slider
    feedList.querySelectorAll('.dailyfeed-oracle-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const taskId = parseInt(btn.dataset.taskId);
        if (taskId) {
          markTaskAsRead(taskId);
          const item = btn.closest('.dailyfeed-item');
          if (item) item.classList.remove('unread');
          showFeedOracleSlider(taskId);
        }
      });
    });

    feedList.querySelectorAll('.dailyfeed-item-title').forEach(title => {
      title.addEventListener('click', () => {
        const item = title.closest('.dailyfeed-item');
        const taskId = item?.dataset?.taskId;
        if (taskId) {
          markTaskAsRead(parseInt(taskId));
          item.classList.remove('unread');
          // Check if task has discussion_summary — open slider with it
          const todo = allTodos.find(t => t.id == taskId);
          if (todo && todo.discussion_summary) {
            showFeedDiscussionSlider(parseInt(taskId));
            return;
          }
        }
        const source = item?.querySelector('.dailyfeed-item-source');
        if (source) window.open(source.href, '_blank');
      });
    });

    // Also handle click on desc body to open slider
    feedList.querySelectorAll('.dailyfeed-item-desc').forEach(desc => {
      desc.style.cursor = 'pointer';
      desc.addEventListener('click', (e) => {
        if (e.target.classList.contains('dailyfeed-view-more')) return; // let view more handle itself
        const item = desc.closest('.dailyfeed-item');
        const taskId = item?.dataset?.taskId;
        if (taskId) {
          markTaskAsRead(parseInt(taskId));
          item.classList.remove('unread');
          const todo = allTodos.find(t => t.id == taskId);
          if (todo && todo.discussion_summary) {
            showFeedDiscussionSlider(parseInt(taskId));
            return;
          }
        }
        const source = item?.querySelector('.dailyfeed-item-source');
        if (source) window.open(source.href, '_blank');
      });
    });
  }

  // === Feed Discussion Summary Slider — shows discussion_summary in a read-only slider ===
  function showFeedDiscussionSlider(taskId) {
    const todo = allTodos.find(t => t.id == taskId);
    if (!todo || !todo.discussion_summary) return;

    const isDarkMode = document.body.classList.contains('dark-mode');

    // Remove any existing discussion or transcript sliders
    document.querySelectorAll('.discussion-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-overlay:not(.discussion-slider-overlay)').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

    const col3Rect = window.Oracle.getCol3Rect();
    const overlay = document.createElement('div');
    overlay.className = 'transcript-slider-overlay discussion-slider-overlay';
    if (col3Rect) {
      overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 9999; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;`;
    } else {
      overlay.style.cssText = 'position: fixed; top: 0; right: 0; bottom: 0; width: 400px; z-index: 9999; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;';
    }

    const slider = document.createElement('div');
    slider.className = 'transcript-slider';
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; overflow: hidden; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 14px 16px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; gap: 10px; flex-shrink: 0; background: ${isDarkMode ? 'rgba(102,126,234,0.12)' : 'rgba(102,126,234,0.06)'};`;
    header.innerHTML = `
      <div style="width:32px;height:32px;min-width:32px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:14px;">💬</div>
      <div style="flex:1;min-width:0;">
        <div style="font-weight:600;font-size:14px;color:${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Discussion Summary</div>
      </div>
      <button class="feed-oracle-close" style="background:transparent;border:none;color:${isDarkMode ? '#888' : '#95a5a6'};cursor:pointer;font-size:20px;padding:4px 8px;border-radius:6px;transition:all 0.2s;">×</button>
    `;
    slider.appendChild(header);

    // Title snippet
    const snippetText = todo.task_title || todo.task_name || '';
    const sourceUrl = todo.message_link || '';
    if (snippetText) {
      const snippet = document.createElement('div');
      snippet.style.cssText = `padding: 10px 16px; font-size: 13px; font-weight: 600; color: ${isDarkMode ? '#ccc' : '#2c3e50'}; line-height: 1.5; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.06)' : 'rgba(225,232,237,0.4)'}; flex-shrink: 0;`;
      let snippetHtml = escapeHtml(snippetText.substring(0, 200)) + (snippetText.length > 200 ? '...' : '');
      if (sourceUrl) {
        let displayDomain = '';
        const isSlack = sourceUrl.includes('slack.com');
        try { displayDomain = new URL(sourceUrl).hostname.replace('www.', ''); } catch(e) { displayDomain = sourceUrl.substring(0, 40); }
        const slackImg = isSlack ? `<img src="${chrome.runtime.getURL('icon-slack.png')}" alt="Slack" style="width:13px;height:13px;object-fit:contain;vertical-align:middle;margin-right:3px;">` : '';
        snippetHtml += `<div style="margin-top:6px;"><a href="${escapeHtml(sourceUrl)}" target="_blank" style="color:#667eea;text-decoration:none;font-size:11px;display:inline-flex;align-items:center;gap:3px;">${slackImg}${isSlack ? 'Open in Slack' : escapeHtml(displayDomain)}</a></div>`;
      }
      snippet.innerHTML = snippetHtml;
      slider.appendChild(snippet);
    }

    // Discussion summary content
    const contentArea = document.createElement('div');
    contentArea.style.cssText = `flex: 1; overflow-y: auto; padding: 16px; font-size: 13.5px; line-height: 1.7; color: ${isDarkMode ? '#d0d0d0' : '#34495e'};`;

    // Format the discussion_summary text
    const taskChipSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0;"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>`;
    let formatted = escapeHtml(todo.discussion_summary);
    // Convert [Task:id] to clickable task chips
    const slackChipImg = `<img src="${chrome.runtime.getURL('icon-slack.png')}" alt="Slack" style="width:13px;height:13px;object-fit:contain;">`;
    formatted = formatted.replace(/\[Task:(\d+)\]/g, (match, taskId) => {
      return `<span class="oracle-task-chip" data-task-id="${taskId}" title="Open task #${taskId}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;background:rgba(102,126,234,0.12);color:#667eea;border-radius:6px;cursor:pointer;vertical-align:middle;transition:all 0.15s;border:1px solid rgba(102,126,234,0.2);margin:0 2px;">${slackChipImg}</span>`;
    });
    formatted = formatted.replace(/\\n/g, '<br>').replace(/\n/g, '<br>');
    // Render markdown bold **text** as <strong>
    formatted = formatted.replace(/\*\*(.+?)\*\*/g, '<strong style="color:' + (isDarkMode ? '#e8e8e8' : '#2c3e50') + ';font-size:14px;display:inline-block;margin-top:8px;">$1</strong>');
    formatted = formatted.replace(/^• /gm, '<span style="color:#667eea;margin-right:4px;">•</span> ');
    formatted = formatted.replace(/^- /gm, '<span style="color:#667eea;margin-right:4px;">•</span> ');
    contentArea.innerHTML = formatted;
    slider.appendChild(contentArea);

    // Bind task chip click handlers
    contentArea.querySelectorAll('.oracle-task-chip').forEach(chip => {
      chip.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const chipTaskId = chip.dataset.taskId;
        if (!chipTaskId) return;
        // Suspend the discussion escape handler while transcript is open on top
        document.removeEventListener('keydown', escHandler);

        // Listen for transcript slider removal (when it closes via Escape)
        const observer = new MutationObserver((mutations) => {
          for (const mutation of mutations) {
            for (const node of mutation.removedNodes) {
              if (node.classList && node.classList.contains('transcript-slider-overlay') && !node.classList.contains('discussion-slider-overlay')) {
                // Transcript closed — re-enable discussion escape handler
                document.addEventListener('keydown', escHandler);
                observer.disconnect();
                return;
              }
            }
          }
        });
        observer.observe(document.body, { childList: true });

        // Open transcript slider on top (z-index 10000 > discussion's 9999)
        if (typeof window.openTaskFromChat === 'function') {
          window.openTaskFromChat(chipTaskId);
        }
      });
    });

    // Bottom bar with actions
    const bottomBar = document.createElement('div');
    bottomBar.style.cssText = `padding: 12px 16px; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; flex-shrink: 0; display: flex; gap: 8px; align-items: center;`;
    bottomBar.innerHTML = `
      <button class="discussion-copy-btn" title="Copy summary" style="background:${isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border:1px solid rgba(102,126,234,0.2);color:#667eea;height:34px;padding:0 12px;border-radius:8px;font-size:12px;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all 0.2s;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>Copy</button>
      <button class="discussion-oracle-btn" data-task-id="${taskId}" title="Ask Oracle about this" style="background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;height:34px;padding:0 12px;border-radius:8px;font-size:12px;cursor:pointer;display:flex;align-items:center;gap:5px;transition:all 0.2s;">∞ Ask Oracle</button>
      <div style="flex:1;"></div>
      <button class="discussion-done-btn" title="Mark as Done" style="background:rgba(39,174,96,0.12);border:1px solid rgba(39,174,96,0.3);color:#27ae60;width:34px;height:34px;border-radius:8px;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;">✓</button>
    `;
    slider.appendChild(bottomBar);

    overlay.appendChild(slider);
    document.body.appendChild(overlay);

    // Close handler
    slider.querySelector('.feed-oracle-close').addEventListener('click', () => {
      overlay.style.animation = 'fadeOut 0.15s ease-out';
      slider.style.animation = 'slideOutRight 0.2s ease-out';
      setTimeout(() => overlay.remove(), 200);
    });

    // Copy handler
    slider.querySelector('.discussion-copy-btn').addEventListener('click', () => {
      const btn = slider.querySelector('.discussion-copy-btn');
      navigator.clipboard.writeText(todo.discussion_summary).then(() => {
        btn.innerHTML = '✓ Copied';
        btn.style.color = '#27ae60';
        setTimeout(() => {
          btn.innerHTML = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>Copy';
          btn.style.color = '#667eea';
        }, 1500);
      });
    });

    // Ask Oracle handler — close this slider and open the Oracle AI slider
    slider.querySelector('.discussion-oracle-btn').addEventListener('click', () => {
      overlay.remove();
      showFeedOracleSlider(taskId);
    });

    // Mark as Done handler
    slider.querySelector('.discussion-done-btn').addEventListener('click', async () => {
      const doneBtn = slider.querySelector('.discussion-done-btn');
      doneBtn.style.background = 'rgba(39,174,96,0.3)';
      doneBtn.textContent = '✓';
      doneBtn.disabled = true;
      allTodos = allTodos.filter(t => t.id != taskId);
      const feedItem = document.querySelector(`.dailyfeed-item[data-task-id="${taskId}"]`);
      if (feedItem) {
        feedItem.classList.add('completing');
        setTimeout(() => {
          feedItem.remove();
          document.querySelectorAll('.dailyfeed-date-group').forEach(group => {
            if (group.querySelectorAll('.dailyfeed-item').length === 0) group.remove();
          });
          if (typeof window.updateDailyFeedCount === 'function') window.updateDailyFeedCount();
        }, 400);
      }
      showToastNotification('Feed item marked as done');
      overlay.style.animation = 'fadeOut 0.15s ease-out';
      slider.style.animation = 'slideOutRight 0.2s ease-out';
      setTimeout(() => overlay.remove(), 200);
      updateTodoField(taskId, 'status', 1).catch(err => console.error('Error marking feed item done:', err));
    });

    // Escape key to close discussion slider with animation
    const escHandler = (e) => {
      if (e.key === 'Escape') {
        e.stopImmediatePropagation();
        overlay.style.animation = 'fadeOut 0.15s ease-out';
        slider.style.animation = 'slideOutRight 0.2s ease-out';
        setTimeout(() => overlay.remove(), 200);
        document.removeEventListener('keydown', escHandler);
      }
    };
    document.addEventListener('keydown', escHandler);
  }

  // === Feed Oracle Slider — a dedicated, lightweight slider for feed items ===
  function showFeedOracleSlider(taskId) {
    const todo = allTodos.find(t => t.id == taskId);
    if (!todo) return;

    const isDarkMode = document.body.classList.contains('dark-mode');

    // Remove any existing sliders
    document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

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
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; overflow: hidden; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 14px 16px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; gap: 10px; flex-shrink: 0; background: ${isDarkMode ? 'rgba(102,126,234,0.12)' : 'rgba(102,126,234,0.06)'};`;
    header.innerHTML = `
      <div style="width:32px;height:32px;min-width:32px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:700;font-size:14px;">∞</div>
      <div style="flex:1;min-width:0;">
        <div style="font-weight:600;font-size:14px;color:${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</div>
      </div>
      <button class="feed-oracle-close" style="background:transparent;border:none;color:${isDarkMode ? '#888' : '#95a5a6'};cursor:pointer;font-size:20px;padding:4px 8px;border-radius:6px;transition:all 0.2s;">×</button>
    `;
    slider.appendChild(header);

    // Context snippet — show feed item title and source URL
    const snippetText = todo.task_title || todo.task_name || '';
    const sourceUrl = todo.message_link || '';
    if (snippetText) {
      const snippet = document.createElement('div');
      snippet.style.cssText = `padding: 10px 16px; font-size: 12px; color: ${isDarkMode ? '#999' : '#7f8c8d'}; line-height: 1.5; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.06)' : 'rgba(225,232,237,0.4)'}; flex-shrink: 0; max-height: 80px; overflow: hidden;`;
      const displayText = snippetText.substring(0, 180);
      let snippetHtml = escapeHtml(displayText) + (snippetText.length > 180 ? '...' : '');
      if (sourceUrl) {
        let domain = '';
        try { domain = new URL(sourceUrl).hostname.replace('www.', ''); } catch(e) { domain = sourceUrl.substring(0, 40); }
        snippetHtml += `<div style="margin-top:4px;"><a href="${escapeHtml(sourceUrl)}" target="_blank" style="color:#667eea;text-decoration:none;font-size:11px;display:inline-flex;align-items:center;gap:4px;"><svg width="11" height="11" viewBox="0 0 24 24" fill="none" style="flex-shrink:0;"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><polyline points="15 3 21 3 21 9" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="10" y1="14" x2="21" y2="3" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>${escapeHtml(domain)}</a></div>`;
      }
      snippet.innerHTML = snippetHtml;
      slider.appendChild(snippet);
    }

    // Response content area
    const contentArea = document.createElement('div');
    contentArea.style.cssText = `flex: 1; overflow-y: auto; padding: 16px; font-size: 14px; line-height: 1.6; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; display: flex; flex-direction: column;`;
    slider.appendChild(contentArea);

    // Reply section
    const replySection = document.createElement('div');
    replySection.style.cssText = `padding: 12px 16px; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; flex-shrink: 0;`;
    replySection.innerHTML = `
      <div style="display:flex;gap:8px;align-items:flex-end;">
        <div class="feed-oracle-reply" contenteditable="true" placeholder="Ask a follow-up..." style="flex:1;padding:10px 14px;border:2px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:10px;font-family:inherit;font-size:13px;min-height:40px;max-height:120px;overflow-y:auto;outline:none;background:${isDarkMode ? '#16213e' : 'white'};color:${isDarkMode ? '#e8e8e8' : '#2c3e50'};line-height:1.4;"></div>
        <button class="feed-oracle-send" style="background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;width:34px;height:34px;border-radius:10px;font-size:14px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;flex-shrink:0;">↗</button>
        <button class="feed-oracle-done" title="Mark as Done" style="background:rgba(39,174,96,0.12);border:1px solid rgba(39,174,96,0.3);color:#27ae60;width:34px;height:34px;border-radius:10px;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;flex-shrink:0;">✓</button>
      </div>
      <div style="margin-top:6px;font-size:11px;color:${isDarkMode ? '#555' : '#bdc3c7'};">Press ⌘+Enter to send</div>
    `;
    slider.appendChild(replySection);

    overlay.appendChild(slider);
    document.body.appendChild(overlay);

    // Close handler
    slider.querySelector('.feed-oracle-close').addEventListener('click', () => {
      overlay.style.animation = 'fadeOut 0.15s ease-out';
      slider.style.animation = 'slideOutRight 0.2s ease-out';
      setTimeout(() => overlay.remove(), 200);
    });

    // Mark as Done handler
    slider.querySelector('.feed-oracle-done').addEventListener('click', async () => {
      const doneBtn = slider.querySelector('.feed-oracle-done');
      doneBtn.style.background = 'rgba(39,174,96,0.3)';
      doneBtn.textContent = '✓';
      doneBtn.disabled = true;
      // Remove from local arrays
      allTodos = allTodos.filter(t => t.id != taskId);
      // Remove the feed item from DOM
      const feedItem = document.querySelector(`.dailyfeed-item[data-task-id="${taskId}"]`);
      if (feedItem) {
        feedItem.classList.add('completing');
        setTimeout(() => {
          feedItem.remove();
          // Clean up empty date groups
          document.querySelectorAll('.dailyfeed-date-group').forEach(group => {
            if (group.querySelectorAll('.dailyfeed-item').length === 0) group.remove();
          });
          if (typeof window.updateDailyFeedCount === 'function') window.updateDailyFeedCount();
        }, 400);
      }
      showToastNotification('Feed item marked as done');
      // Close slider
      overlay.style.animation = 'fadeOut 0.15s ease-out';
      slider.style.animation = 'slideOutRight 0.2s ease-out';
      setTimeout(() => overlay.remove(), 200);
      // Send to backend
      updateTodoField(taskId, 'status', 1).catch(err => console.error('Error marking feed item done:', err));
    });

    // Escape key to close
    const escHandler = (e) => {
      if (e.key === 'Escape') {
        overlay.remove();
        document.removeEventListener('keydown', escHandler);
      }
    };
    document.addEventListener('keydown', escHandler);

    // Conversation state
    let oracleConversationHistory = [];
    let oracleSessionId = 'feed_' + taskId + '_' + Date.now();

    // Format response helper
    const formatResponse = (text) => {
      let formatted = escapeHtml(text);
      const bgColor = isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)';
      const urlPattern = /\[(https?:\/\/[^\]]+)\]|(?<!\[)(https?:\/\/[^\s\]<>"]+)/g;
      formatted = formatted.replace(urlPattern, (match, bracketUrl, plainUrl) => {
        const url = bracketUrl || plainUrl;
        let icon = '🔗';
        if (url.includes('slack.com')) icon = '💬';
        else if (url.includes('docs.google.com')) icon = '📄';
        else if (url.includes('drive.google.com')) icon = '📁';
        else if (url.includes('github.com')) icon = '🐙';
        return '<a href="' + url + '" target="_blank" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;background:' + bgColor + ';border-radius:4px;text-decoration:none;margin:0 2px;border:1px solid rgba(102,126,234,0.2);vertical-align:middle;">' + icon + '</a>';
      });
      formatted = formatted.replace(/\\n/g, '<br>').replace(/\n/g, '<br>');
      formatted = formatted.replace(/^- /gm, '• ');
      return formatted;
    };

    // Copy buttons helper
    const addCopyButtons = (rawText) => {
      const existing = contentArea.querySelector('.oracle-copy-actions');
      if (existing) existing.remove();
      const copySvg = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';
      const actionsDiv = document.createElement('div');
      actionsDiv.className = 'oracle-copy-actions';
      actionsDiv.style.cssText = 'display:flex;gap:6px;margin-top:12px;padding-top:8px;border-top:1px solid ' + (isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(0,0,0,0.06)') + ';';
      const btn = document.createElement('button');
      btn.title = 'Copy message';
      btn.innerHTML = copySvg;
      btn.style.cssText = 'background:' + (isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)') + ';border:1px solid rgba(102,126,234,0.2);color:#667eea;width:28px;height:28px;border-radius:6px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;';
      btn.addEventListener('click', () => {
        navigator.clipboard.writeText(rawText).then(() => {
          btn.innerHTML = '✓'; btn.style.color = '#27ae60';
          setTimeout(() => { btn.innerHTML = copySvg; btn.style.color = '#667eea'; }, 1500);
        });
      });
      actionsDiv.appendChild(btn);
      const responseDiv = contentArea.querySelector('.oracle-response:last-of-type');
      if (responseDiv) responseDiv.appendChild(actionsDiv);
    };

    // Core send function
    const sendMessage = async (userMessage, isFollowUp) => {
      if (!isFollowUp) {
        contentArea.innerHTML = '';
      } else {
        const userBubble = document.createElement('div');
        userBubble.style.cssText = 'padding:8px 12px;background:' + (isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)') + ';border-radius:10px;margin-bottom:12px;margin-left:auto;font-size:13px;color:' + (isDarkMode ? '#e0e0e0' : '#2c3e50') + ';max-width:85%;text-align:right;';
        userBubble.textContent = userMessage;
        contentArea.appendChild(userBubble);
      }

      // Show thinking indicator as a left-aligned bubble
      const loadingDiv = document.createElement('div');
      loadingDiv.className = 'oracle-loading';
      loadingDiv.style.cssText = 'margin-right:auto;max-width:85%;';
      loadingDiv.innerHTML = '<div style="padding:10px 14px;background:' + (isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)') + ';border-radius:14px 14px 14px 4px;font-size:13px;color:' + (isDarkMode ? '#888' : '#7f8c8d') + ';"><span class="typing-dots">Thinking<span>.</span><span>.</span><span>.</span></span></div>';
      contentArea.appendChild(loadingDiv);
      contentArea.scrollTop = contentArea.scrollHeight;

      let userId = window.Oracle?.state?.userData?.userId;
      if (!userId) {
        try { const r = await chrome.storage.local.get(['oracle_user_data']); userId = r?.oracle_user_data?.userId || null; } catch { userId = null; }
      }

      oracleConversationHistory.push({ role: 'user', content: userMessage });

      const payload = {
        message: userMessage,
        session_id: oracleSessionId,
        conversation: oracleConversationHistory,
        timestamp: new Date().toISOString(),
        source: isFollowUp ? 'oracle-feed-followup' : 'oracle-feed',
        user_id: userId,
        context: {
          task_title: todo.task_title || '',
          task_name: todo.task_name || '',
          message_link: todo.message_link || ''
        }
      };

      try {
        let fullResponseText = '', streamStarted = false;
        const controller = new AbortController();
        const streamTimeout = setTimeout(() => {
          controller.abort();
          if (!streamStarted) {
            const loadEl = contentArea.querySelector('.oracle-loading');
            if (loadEl) loadEl.innerHTML = '<div style="padding:16px;color:#e74c3c;text-align:center;"><span style="font-size:24px;">⏱</span><div>Response timed out</div></div>';
          }
        }, 120000);

        const response = await fetch('https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
          signal: controller.signal
        });

        if (!response.ok) throw new Error('HTTP ' + response.status);

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        try {
          while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';
            for (const line of lines) {
              if (!line.trim()) continue;
              try {
                const parsed = JSON.parse(line);
                if (parsed.type === 'item' && parsed.content && parsed.metadata?.nodeName !== 'Respond to Webhook') {
                  if (!streamStarted) {
                    streamStarted = true;
                    const loadEl = contentArea.querySelector('.oracle-loading');
                    if (loadEl) loadEl.remove();
                    const responseDiv = document.createElement('div');
                    responseDiv.className = 'oracle-response';
                    responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
                    contentArea.appendChild(responseDiv);
                  }
                  fullResponseText += parsed.content;
                  const responseDiv = contentArea.querySelector('.oracle-response:last-of-type');
                  if (responseDiv) {
                    responseDiv.innerHTML = (window.OracleAssistant?.formatChatResponseWithAnnotations || formatResponse)(fullResponseText);
                    if (window.OracleAssistant?.bindTaskChips) window.OracleAssistant.bindTaskChips(responseDiv);
                  }
                  contentArea.scrollTop = contentArea.scrollHeight;
                }
              } catch { /* skip */ }
            }
          }
          if (buffer.trim()) {
            try {
              const parsed = JSON.parse(buffer);
              if (parsed.type === 'item' && parsed.content && parsed.metadata?.nodeName !== 'Respond to Webhook') {
                fullResponseText += parsed.content;
              }
            } catch { /* skip */ }
          }
        } finally {
          clearTimeout(streamTimeout);
          reader.releaseLock();
        }

        if (!streamStarted) {
          const loadEl = contentArea.querySelector('.oracle-loading');
          if (loadEl) loadEl.remove();
          const responseDiv = document.createElement('div');
          responseDiv.className = 'oracle-response';
          responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
          contentArea.appendChild(responseDiv);
        }

        oracleConversationHistory.push({ role: 'assistant', content: fullResponseText });
        const finalDiv = contentArea.querySelector('.oracle-response:last-of-type');
        if (finalDiv) {
          finalDiv.innerHTML = (window.OracleAssistant?.formatChatResponseWithAnnotations || formatResponse)(fullResponseText);
          if (window.OracleAssistant?.bindTaskChips) window.OracleAssistant.bindTaskChips(finalDiv);
        }
        addCopyButtons(fullResponseText);
        contentArea.scrollTop = contentArea.scrollHeight;
      } catch (error) {
        console.error('Feed Oracle request failed:', error);
        const loadEl = contentArea.querySelector('.oracle-loading');
        if (loadEl) loadEl.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;gap:12px;color:#e74c3c;"><span style="font-size:24px;">⚠️</span><div>Failed to get response</div><div style="font-size:12px;color:' + (isDarkMode ? '#888' : '#95a5a6') + ';">' + escapeHtml(error.message) + '</div></div>';
      }
    };

    // Send button & keyboard shortcut
    const sendBtn = slider.querySelector('.feed-oracle-send');
    const replyInput = slider.querySelector('.feed-oracle-reply');

    const handleSend = () => {
      const text = replyInput.textContent.trim();
      if (!text) return;
      replyInput.textContent = '';
      sendMessage(text, true);
    };

    sendBtn.addEventListener('click', handleSend);
    replyInput.addEventListener('keydown', (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        handleSend();
      }
    });

    // Auto-trigger initial analysis
    const feedQuery = todo.task_title || todo.task_name || 'Feed item';
    const feedUrl = todo.message_link || '';
    const feedMsg = 'Analyze this feed item and provide key insights and takeaways - ' + feedQuery + (feedUrl ? '. Get further details for URL - ' + feedUrl : '');
    sendMessage(feedMsg, false);
  }

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
      showToastNotification(message);
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
    const actionUnread = (typeof pendingActionUpdates !== 'undefined' ? pendingActionUpdates.length : 0);
    const fyiUnread = (typeof pendingFyiUpdates !== 'undefined' ? pendingFyiUpdates.length : 0);
    const total = actionUnread + fyiUnread;
    try {
      chrome.runtime.sendMessage({ type: 'updateBadge', actionUnread, fyiUnread, total });
    } catch (e) {}
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
    const nonMeetingTodos = todos.filter(t => !isMeetingLink(t.message_link) && !(t.tags || []).some(tag => String(tag) === '2282' || String(tag).toLowerCase() === 'daily feed'));

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
    const typeIconHtml = getTypeIconHtml(todo.type || null, todo.participant_text || '');

    // # icon for Slack tasks (ungrouped) - only show for slack_thread type
    const slackChannelId = extractSlackChannelId(messageLink);
    const channelName = todo.participant_text || '';
    const eyeIconHtml = (slackChannelId && todo.type === 'slack_thread')
      ? `<span class="todo-slack-eye-btn" data-channel-id="${escapeHtml(slackChannelId)}" data-channel-name="${escapeHtml(channelName)}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;">#</span>`
      : '';

    // Check if task is unread
    const isUnread = isTaskUnread(todo.id, todo.updated_at);

    return `<div class="todo-item ${todo.status === 1 ? 'completed' : ''} ${newItemIds.includes(todo.id) ? 'appearing' : ''} ${isUnread ? 'unread' : ''}" style="animation-delay: ${index * 0.05}s" data-todo-id="${todo.id}">
            ${getTrendingIconHtml(todo)}${eyeIconHtml}
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
                ${(typeIconHtml || slackChannelId) ? `<span style="flex:1;min-width:4px;"></span>` : ''}
                ${typeIconHtml || ''}
                ${slackChannelId ? `<span class="slack-bell-channel-btn" data-channel-id="${escapeHtml(slackChannelId)}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-left:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
              </div>
              ${tagsHtml}
            </div>
            <div class="todo-actions">
              ${showClock ? '<div class="todo-clock" data-todo-id="' + todo.id + '" title="Remind me in">🕐</div>' : ''}
            </div>
            ${messageLink && messageLink.includes('mail.google.com') ? `<span class="todo-block-email-btn" data-todo-id="${todo.id}" title="Block email participants"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
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

    // Deduplicate meetings by title + due_by + participant_text
    const seen = new Set();
    const dedupedMeetings = meetings.filter(m => {
      const key = `${(m.task_title || m.task_name || '').trim().toLowerCase()}|${m.due_by || ''}|${(m.participant_text || '').replace(/\|self_accepted:(true|false)/, '').replace(/\|\d+\s*(?:mins?|minutes?|hrs?|hours?)/i, '').trim().toLowerCase()}`;
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });

    // Sort meetings by due_by
    const sortedMeetings = [...dedupedMeetings].sort((a, b) => {
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

    // Helper to check if a meeting is active (not cancelled, not overdue/past)
    const isActiveMeeting = (m) => {
      const title = (m.task_title || m.task_name || '').toLowerCase();
      if (title.startsWith('cancelled') || title.startsWith('canceled')) return false;
      if (m.due_by && new Date(m.due_by) < new Date()) {
        // Check if currently underway (buildMeetingItemHtml sets _isCurrentlyUnderway but hasn't run yet here)
        const pt = m.participant_text || '';
        const durMatch = pt.match(/\|(\d+)\s*(?:mins?|minutes?)/i) || pt.match(/\|(\d+)\s*(?:hrs?|hours?)/i);
        let dur = durMatch ? parseInt(durMatch[1]) : 30;
        if (durMatch && pt.match(/\|(\d+)\s*(?:hrs?|hours?)/i) && !pt.match(/\|(\d+)\s*(?:mins?|minutes?)/i)) dur *= 60;
        const start = new Date(m.due_by);
        const end = new Date(start.getTime() + dur * 60000);
        if (end > new Date()) return true; // currently underway
        return false; // truly overdue
      }
      return true;
    };

    // Get sorted date keys — only include dates that have at least one active meeting
    const todayKey = getDateKey(new Date().toISOString());
    let dateKeys = Object.keys(meetingsByDate).filter(key => key !== 'no-date' && key >= todayKey).sort();
    dateKeys = dateKeys.filter(dk => meetingsByDate[dk].some(m => isActiveMeeting(m)));
    if (meetingsByDate['no-date'] && meetingsByDate['no-date'].some(m => isActiveMeeting(m))) {
      dateKeys.push('no-date');
    }

    // Check for meaningful meeting changes (new, rescheduled, cancelled)
    const currentMeetingIds = new Set(meetings.map(m => String(m.id)));
    const newOrModifiedMeetingIds = new Set();

    // Build snapshot of meaningful fields: title, description, participant_text, due_by
    // Ignores updated_at and status which change on every sync
    const currentMeetingSnapshots = new Map();
    meetings.forEach(m => {
      currentMeetingSnapshots.set(String(m.id), `${(m.task_title || m.task_name || '').trim()}|${(m.description || '').trim()}|${(m.participant_text || '').replace(/\|self_accepted:(true|false)/, '').replace(/\|\d+\s*(?:mins?|minutes?|hrs?|hours?)/i, '').trim()}|${m.due_by || ''}`);
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
            // Title, description, participants, or time changed
            console.log('Meeting changed (title/desc/participants/time):', id, 'was:', prevSnapshot, 'now:', currSnapshot);
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
          <span class="meeting-date-count">${meetingsOnDate.filter(m => !(m.task_title || m.task_name || '').toLowerCase().startsWith('cancelled')).length}</span>
        </div>
      `;
    }).join('');

    // Build meeting content sections for each date
    const isCancelledMeeting = (m) => {
      const title = (m.task_title || m.task_name || '').toLowerCase();
      return title.startsWith('cancelled') || title.startsWith('canceled');
    };

    const dateContentSections = dateKeys.map((dateKey, index) => {
      const meetingsOnDate = meetingsByDate[dateKey];
      const isActive = index === 0;

      // Separate active vs cancelled
      const activeMeetings = meetingsOnDate.filter(m => !isCancelledMeeting(m));
      const cancelledMeetings = meetingsOnDate.filter(m => isCancelledMeeting(m));

      // Sort active meetings: unread first, then by time
      const sortedActive = [...activeMeetings].sort((a, b) => {
        const aUnread = newOrModifiedMeetingIds.has(String(a.id));
        const bUnread = newOrModifiedMeetingIds.has(String(b.id));
        if (aUnread && !bUnread) return -1;
        if (!aUnread && bUnread) return 1;
        if (a.due_by && b.due_by) return new Date(a.due_by) - new Date(b.due_by);
        return 0;
      });

      const sortedCancelled = [...cancelledMeetings].sort((a, b) => {
        if (a.due_by && b.due_by) return new Date(a.due_by) - new Date(b.due_by);
        return 0;
      });

      const activeItems = sortedActive.map(meeting => {
        const isUnread = newOrModifiedMeetingIds.has(String(meeting.id));
        return buildMeetingItemHtml(meeting, googleCalendarIconUrl, zoomIconUrl, googleMeetIconUrl, isUnread);
      }).join('');

      let cancelledHtml = '';
      if (sortedCancelled.length > 0) {
        const cancelledItems = sortedCancelled.map(meeting => {
          const isUnread = newOrModifiedMeetingIds.has(String(meeting.id));
          return buildMeetingItemHtml(meeting, googleCalendarIconUrl, zoomIconUrl, googleMeetIconUrl, isUnread);
        }).join('');
        cancelledHtml = `
          <div class="cancelled-meetings-sub-accordion" style="margin-top: 8px; border-top: 1px solid var(--border-light, rgba(225,232,237,0.3));">
            <div class="cancelled-sub-header" style="display: flex; align-items: center; gap: 6px; padding: 8px 12px; cursor: pointer; user-select: none; opacity: 0.7; transition: opacity 0.2s;">
              <span class="cancelled-chevron" style="font-size: 10px; transition: transform 0.2s;">▶</span>
              <span style="font-size: 12px; font-weight: 500; color: var(--text-muted, #95a5a6);">Cancelled (${sortedCancelled.length})</span>
            </div>
            <div class="cancelled-sub-content" style="display: none;">
              ${cancelledItems}
            </div>
          </div>
        `;
      }

      return `
        <div class="meetings-date-content ${isActive ? 'active' : ''}" data-date-key="${dateKey}">
          <div class="meetings-list">
            ${activeItems}
          </div>
          ${cancelledHtml}
        </div>
      `;
    }).join('');

    // If no future dates remain, don't render the accordion
    const visibleMeetingCount = dateKeys.reduce((sum, dk) => sum + meetingsByDate[dk].filter(m => isActiveMeeting(m)).length, 0);
    if (!dateKeys.length) return '';

    // Find currently underway meeting for header display (exclude cancelled)
    const underwayMeeting = sortedMeetings.find(m => m._isCurrentlyUnderway && !isCancelledMeeting(m));
    const underwayJoinLink = underwayMeeting?._conferencingLink || underwayMeeting?._primaryLink || '';
    const underwayHtml = (underwayMeeting && underwayJoinLink) ? `
      <a href="${escapeHtml(underwayJoinLink)}" target="_blank" class="meeting-underway-join" title="${escapeHtml((underwayMeeting.task_title || 'Meeting') + ' — Click to join')}" style="display:flex;align-items:center;gap:6px;padding:4px 10px;background:rgba(39,174,96,0.12);border:1px solid rgba(39,174,96,0.3);border-radius:8px;color:#27ae60;text-decoration:none;font-size:11px;font-weight:600;transition:all 0.2s;white-space:nowrap;max-width:240px;overflow:hidden;">
        <span style="font-size:8px;">🟢</span>
        <span style="overflow:hidden;text-overflow:ellipsis;">${escapeHtml((underwayMeeting.task_title || 'Meeting').length > 30 ? (underwayMeeting.task_title || 'Meeting').substring(0, 28) + '...' : (underwayMeeting.task_title || 'Meeting'))}</span>
        <span style="font-size:10px;">↗</span>
      </a>` : '';

    return `
      <div class="meetings-accordion ${hasAnyUpdates ? 'has-updates' : ''}">
        <div class="meetings-accordion-header">
          <div class="meetings-accordion-title">
            <span>📅</span>
            <span>Meetings</span>
            <span class="meetings-accordion-count">${visibleMeetingCount}</span>
          </div>
          <div style="display:flex;align-items:center;gap:8px;">
            ${underwayHtml}
            <span class="meetings-accordion-chevron">▼</span>
          </div>
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
      if (e.key === 'Escape') { if (document.querySelector('.attachment-preview-modal')) return; e.stopImmediatePropagation(); closeSlider(); document.removeEventListener('keydown', escHandler); }
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
                ${p.comment ? `<div style="font-size: 11px; color: ${isDarkMode ? '#a0a0a0' : '#6c7a89'}; margin-top: 3px; font-style: italic; white-space: normal; word-break: break-word;">💬 ${escapeHtml(p.comment)}</div>` : ''}
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

      // 5. Description
      const description = meetingDetails.description || '';
      if (description) {
        // Strip HTML tags for clean text display, but preserve line breaks
        const cleanDesc = description
          .replace(/<br\s*\/?>/gi, '\n')
          .replace(/<\/p>/gi, '\n')
          .replace(/<[^>]+>/g, '')
          .replace(/&nbsp;/gi, ' ')
          .replace(/&amp;/gi, '&')
          .replace(/&lt;/gi, '<')
          .replace(/&gt;/gi, '>')
          .replace(/\n{3,}/g, '\n\n')
          .trim();
        if (cleanDesc) {
          // Auto-hyperlink URLs in the plain text description
          const linkedDesc = escapeHtml(cleanDesc).replace(
            /(https?:\/\/[^\s<>"')\]]+)/gi,
            '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$1</a>'
          );
          detailHtml += `
            <div style="display: flex; align-items: flex-start; gap: 12px;">
              <div style="width: 36px; height: 36px; background: rgba(52,152,219,0.1); border-radius: 8px; display: flex; align-items: center; justify-content: center; font-size: 18px; flex-shrink: 0;">📝</div>
              <div style="flex: 1; min-width: 0;">
                <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; margin-bottom: 8px;">Description</div>
                <div style="font-size: 12px; color: ${isDarkMode ? '#a0a0a0' : '#555'}; line-height: 1.6; white-space: pre-wrap; word-break: break-word; ">${linkedDesc}</div>
              </div>
            </div>
          `;
        }
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

    // Extract organiser from participant_text (first name listed) and parse self_accepted flag + duration
    const rawParticipantText = meeting.participant_text || '';
    const selfAcceptedMatch = rawParticipantText.match(/\|self_accepted:(true|false)/);
    const selfAccepted = selfAcceptedMatch ? selfAcceptedMatch[1] === 'true' : null;
    // Parse duration: e.g. "30 mins", "1 hour", "60 mins"
    const durationMatch = rawParticipantText.match(/\|(\d+)\s*(?:mins?|minutes?)/i) || rawParticipantText.match(/\|(\d+)\s*(?:hrs?|hours?)/i);
    let durationMins = 0;
    if (durationMatch) {
      durationMins = parseInt(durationMatch[1]);
      if (rawParticipantText.match(/\|(\d+)\s*(?:hrs?|hours?)/i) && !rawParticipantText.match(/\|(\d+)\s*(?:mins?|minutes?)/i)) durationMins *= 60;
    }
    // Strip the flag and duration from participant_text for display
    const cleanParticipantText = rawParticipantText.replace(/\|self_accepted:(true|false)/g, '').replace(/\|\d+\s*(?:mins?|minutes?|hrs?|hours?)/gi, '');
    const organiserName = extractOrganiser(cleanParticipantText);

    // Determine if meeting is currently underway (started but not ended)
    // Don't show underway for completed meetings (status=1)
    const meetingStart = meeting.due_by ? new Date(meeting.due_by) : null;
    const now = new Date();
    const underwayDurationMins = durationMins > 0 ? durationMins : 30;
    const meetingEnd = meetingStart ? new Date(meetingStart.getTime() + underwayDurationMins * 60000) : null;
    const isCancelled = (meeting.task_title || meeting.task_name || '').toLowerCase().startsWith('cancelled');
    const earlyJoinMs = 60000; // Show as underway 1 minute before start
    const isCurrentlyUnderway = !isCancelled && meeting.status !== 1 && meetingStart && meetingStart.getTime() - earlyJoinMs <= now.getTime() && meetingEnd && meetingEnd > now;
    const isPastMeeting = meetingStart && meetingStart < now && !isCurrentlyUnderway;

    // For underway meetings, show actual start time instead of "Overdue"
    let displayTimeText = timeText;
    if (isCurrentlyUnderway && meetingStart) {
      let h = meetingStart.getHours();
      const m = meetingStart.getMinutes();
      const ap = h >= 12 ? 'PM' : 'AM';
      h = h % 12 || 12;
      const ts = m === 0 ? `${h} ${ap}` : `${h}:${m.toString().padStart(2, '0')} ${ap}`;
      displayTimeText = `Today ${ts}`;
    }
    const timeWithOrganiser = displayTimeText ? (organiserName ? `${displayTimeText} | ${organiserName}` : displayTimeText) : '';
    // Store underway/link info on the meeting object for accordion header access
    meeting._isCurrentlyUnderway = isCurrentlyUnderway;
    meeting._primaryLink = primaryLink;
    // Find the best conferencing link (Zoom/Meet) for the join button, fallback to primary
    const conferencingLink = secondaryLinks.find(l => l && (l.includes('zoom.us') || l.includes('zoom.com') || l.includes('meet.google.com') || l.includes('teams.microsoft.com')));
    meeting._conferencingLink = conferencingLink
      ? ((conferencingLink.includes('zoom.us') || conferencingLink.includes('zoom.com')) ? convertZoomToDeepLink(conferencingLink) : conferencingLink)
      : (primaryLink.includes('zoom.us') || primaryLink.includes('zoom.com') || primaryLink.includes('meet.google.com') || primaryLink.includes('teams.') ? primaryLink : null);
    meeting._primaryTitle = primaryTitle;

    // RSVP icon: ✓ if accepted, ✗ if not accepted, ↕ if unknown
    let rsvpIcon = '↕';
    let rsvpColor = '#667eea';
    if (selfAccepted === true) { rsvpIcon = '✓'; rsvpColor = '#27ae60'; }
    else if (selfAccepted === false) { rsvpIcon = '✗'; rsvpColor = '#e74c3c'; }

    return `
      <div class="meeting-item ${isUnread ? 'unread' : ''} ${isCurrentlyUnderway ? 'underway' : ''}" data-meeting-id="${meeting.id}" data-todo-id="${meeting.id}">
        <div class="meeting-checkbox" data-todo-id="${meeting.id}" title="Mark as done"></div>
        <div class="meeting-info">
          <div class="meeting-title" title="${escapeHtml(title)}">${escapeHtml(truncatedTitle)}</div>
          ${timeWithOrganiser ? `<div class="meeting-time ${isCurrentlyUnderway ? 'underway' : (isPastMeeting ? 'overdue' : '')}">${isCurrentlyUnderway ? '🟢 Currently underway · ' : ''}${escapeHtml(timeWithOrganiser)}${durationMins > 0 ? ` · ${durationMins} mins` : ''}</div>` : ''}
        </div>
        <div class="meeting-actions" style="display: flex; align-items: center; gap: 8px;">
          ${secondaryIconsHtml}
          <a href="${primaryLink}" target="_blank" class="meeting-join-btn meeting-mark-read" title="${primaryTitle}" style="display: flex; align-items: center; justify-content: center; width: 32px; height: 32px; background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.2); border-radius: 6px; transition: all 0.2s;">
            <img src="${primaryIconUrl}" alt="${primaryTitle}" style="width: 18px; height: 18px;">
          </a>
          <div class="meeting-rsvp-btn" title="Accept / Decline" style="position: relative; display: flex; align-items: center; justify-content: center; width: 32px; height: 32px; background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.2); border-radius: 6px; cursor: pointer; transition: all 0.2s; font-size: 14px;">
            <span class="rsvp-icon" style="color: ${rsvpColor}; font-weight: 600;">${rsvpIcon}</span>
          </div>
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
        if (e.target.closest('.meeting-rsvp-btn')) return;
        if (e.target.closest('.rsvp-dropdown')) return;
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

        // Check if Command/Ctrl/Shift key is pressed for multi-select
        if (e.metaKey || e.ctrlKey || e.shiftKey) {
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

    // Setup cancelled meetings sub-accordion toggles
    accordion.querySelectorAll('.cancelled-sub-header').forEach(subHeader => {
      subHeader.addEventListener('click', (e) => {
        e.stopPropagation();
        const subAccordion = subHeader.closest('.cancelled-meetings-sub-accordion');
        const content = subAccordion.querySelector('.cancelled-sub-content');
        const chevron = subHeader.querySelector('.cancelled-chevron');
        const isOpen = content.style.display !== 'none';
        content.style.display = isOpen ? 'none' : 'block';
        chevron.style.transform = isOpen ? 'rotate(0deg)' : 'rotate(90deg)';
        subHeader.style.opacity = isOpen ? '0.7' : '1';
      });
    });

    // Setup accept/decline dropdown buttons
    accordion.querySelectorAll('.meeting-rsvp-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const meetingItem = btn.closest('.meeting-item');
        const meetingId = meetingItem.dataset.meetingId || meetingItem.dataset.todoId;
        
        // Remove any existing dropdown + backdrop
        document.querySelectorAll('.rsvp-dropdown').forEach(d => d.remove());
        document.querySelectorAll('.rsvp-backdrop').forEach(b => b.remove());
        
        const isDarkMode = document.body.classList.contains('dark-mode');

        // Invisible backdrop to catch outside clicks
        const backdrop = document.createElement('div');
        backdrop.className = 'rsvp-backdrop';
        backdrop.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 10000; background: transparent;';
        document.body.appendChild(backdrop);

        const dropdown = document.createElement('div');
        dropdown.className = 'rsvp-dropdown';
        dropdown.style.cssText = `
          position: absolute; top: calc(100% + 6px); right: 0; z-index: 10001;
          background: ${isDarkMode ? '#1e2d4a' : 'white'};
          border: 1px solid ${isDarkMode ? 'rgba(102,126,234,0.25)' : 'rgba(0,0,0,0.08)'};
          border-radius: 10px;
          box-shadow: 0 8px 24px rgba(0,0,0,${isDarkMode ? '0.5' : '0.18'}), 0 2px 8px rgba(0,0,0,${isDarkMode ? '0.3' : '0.08'});
          overflow: hidden; min-width: 140px;
          opacity: 0; transform: translateY(-8px) scale(0.95);
          transition: opacity 0.18s ease-out, transform 0.18s ease-out;
          pointer-events: auto;
        `;
        
        dropdown.innerHTML = `
          <div class="rsvp-option rsvp-accept" style="
            display: flex; align-items: center; gap: 10px; padding: 11px 16px;
            cursor: pointer; font-size: 13px; font-weight: 600; color: #27ae60;
            transition: background 0.15s, transform 0.1s; border-radius: 10px 10px 0 0;
          ">
            <span style="font-size: 15px; display: flex; align-items: center; justify-content: center; width: 22px; height: 22px; background: rgba(39,174,96,0.12); border-radius: 6px;">✓</span>
            Accept
          </div>
          <div style="height: 1px; background: ${isDarkMode ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)'}; margin: 0 10px;"></div>
          <div class="rsvp-option rsvp-decline" style="
            display: flex; align-items: center; gap: 10px; padding: 11px 16px;
            cursor: pointer; font-size: 13px; font-weight: 600; color: #e74c3c;
            transition: background 0.15s, transform 0.1s; border-radius: 0 0 10px 10px;
          ">
            <span style="font-size: 15px; display: flex; align-items: center; justify-content: center; width: 22px; height: 22px; background: rgba(231,76,60,0.12); border-radius: 6px;">✗</span>
            Decline
          </div>
        `;
        
        // Hover states
        dropdown.querySelectorAll('.rsvp-option').forEach(opt => {
          opt.addEventListener('mouseenter', () => {
            opt.style.background = isDarkMode ? 'rgba(102,126,234,0.12)' : 'rgba(0,0,0,0.04)';
            opt.style.transform = 'translateX(2px)';
          });
          opt.addEventListener('mouseleave', () => {
            opt.style.background = 'transparent';
            opt.style.transform = 'translateX(0)';
          });
        });

        // Position relative to button — use fixed positioning on body so z-index works above backdrop
        const btnRect = btn.getBoundingClientRect();
        const dropdownHeight = 90;
        const spaceBelow = window.innerHeight - btnRect.bottom;
        const showAbove = spaceBelow < dropdownHeight + 12;

        if (showAbove) {
          dropdown.style.position = 'fixed';
          dropdown.style.top = 'auto';
          dropdown.style.bottom = (window.innerHeight - btnRect.top + 6) + 'px';
          dropdown.style.right = (window.innerWidth - btnRect.right) + 'px';
          dropdown.style.transform = 'translateY(8px) scale(0.95)';
        } else {
          dropdown.style.position = 'fixed';
          dropdown.style.top = (btnRect.bottom + 6) + 'px';
          dropdown.style.right = (window.innerWidth - btnRect.right) + 'px';
        }

        document.body.appendChild(dropdown);

        // Animate in
        requestAnimationFrame(() => {
          requestAnimationFrame(() => {
            dropdown.style.opacity = '1';
            dropdown.style.transform = 'translateY(0) scale(1)';
          });
        });

        const closeDropdown = () => {
          dropdown.style.opacity = '0';
          dropdown.style.transform = showAbove ? 'translateY(8px) scale(0.95)' : 'translateY(-8px) scale(0.95)';
          backdrop.remove();
          setTimeout(() => dropdown.remove(), 180);
        };
        
        const handleRsvp = async (response) => {
          // Animate the selected option
          const selectedOpt = dropdown.querySelector(response === 'accepted' ? '.rsvp-accept' : '.rsvp-decline');
          if (selectedOpt) {
            selectedOpt.style.transform = 'scale(0.95)';
            setTimeout(() => { selectedOpt.style.transform = 'scale(1)'; }, 100);
          }

          closeDropdown();

          // Visual feedback on the RSVP button
          const icon = btn.querySelector('.rsvp-icon');
          if (icon) {
            icon.style.transition = 'transform 0.2s, color 0.2s';
            icon.style.transform = 'scale(0)';
            setTimeout(() => {
              icon.textContent = response === 'accepted' ? '✓' : '✗';
              icon.style.color = response === 'accepted' ? '#27ae60' : '#e74c3c';
              icon.style.transform = 'scale(1.2)';
              setTimeout(() => { icon.style.transform = 'scale(1)'; }, 150);
            }, 150);
          }

          // Send RSVP to n8n
          const meeting = allCalendarItems.find(t => String(t.id) === String(meetingId));
          console.log('RSVP sending:', response, 'for meeting:', meetingId, 'link:', meeting?.message_link);
          try {
            await fetch(WEBHOOK_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify(createAuthenticatedPayload({
                action: 'calendar_rsvp',
                todo_id: meetingId,
                message_link: meeting ? meeting.message_link : '',
                rsvp_response: response,
                timestamp: new Date().toISOString(),
                source: 'oracle-chrome-extension-newtab'
              }))
            });
          } catch (err) {
            console.error('RSVP error:', err);
          }
        };
        
        dropdown.querySelector('.rsvp-accept').addEventListener('click', (e) => { e.stopPropagation(); handleRsvp('accepted'); });
        dropdown.querySelector('.rsvp-decline').addEventListener('click', (e) => { e.stopPropagation(); handleRsvp('declined'); });

        // Close on backdrop click or Escape
        backdrop.addEventListener('click', closeDropdown);
        const escHandler = (e) => {
          if (e.key === 'Escape') { closeDropdown(); document.removeEventListener('keydown', escHandler); }
        };
        document.addEventListener('keydown', escHandler);
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

    // Sort by latest update (newest first)
    filteredGroups.sort((a, b) => new Date(b.latestUpdate) - new Date(a.latestUpdate));

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
            <button class="document-copy-btn" data-file-url="${fileUrl}" title="Copy document link">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
              </svg>
            </button>
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
          <div style="display: flex; align-items: center; gap: 6px;">
            <button class="documents-accordion-search-btn" title="Search documents">🔍</button>
            <span class="documents-accordion-chevron">▼</span>
          </div>
        </div>
        <div class="documents-search-bar">
          <input type="text" class="documents-search-input" placeholder="Search documents..." maxlength="200">
          <button class="documents-search-close" title="Close search">×</button>
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

    // Search functionality
    const searchBtn = accordion.querySelector('.documents-accordion-search-btn');
    const searchBar = accordion.querySelector('.documents-search-bar');
    const searchInput = accordion.querySelector('.documents-search-input');
    const searchClose = accordion.querySelector('.documents-search-close');

    if (searchBtn && searchBar && searchInput) {
      searchBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        searchBar.classList.toggle('visible');
        if (searchBar.classList.contains('visible')) {
          // Ensure accordion is open when searching
          if (!accordion.classList.contains('open')) {
            accordion.classList.add('open');
          }
          searchInput.focus();
        } else {
          searchInput.value = '';
          // Show all items
          accordion.querySelectorAll('.document-accordion-item').forEach(item => {
            item.style.display = '';
          });
        }
      });

      searchInput.addEventListener('input', () => {
        const query = searchInput.value.toLowerCase().trim();
        accordion.querySelectorAll('.document-accordion-item').forEach(item => {
          const title = item.querySelector('.document-item-title')?.textContent?.toLowerCase() || '';
          const meta = item.querySelector('.document-item-meta')?.textContent?.toLowerCase() || '';
          const taskTexts = Array.from(item.querySelectorAll('.document-task-text')).map(el => el.textContent.toLowerCase()).join(' ');
          const match = !query || title.includes(query) || meta.includes(query) || taskTexts.includes(query);
          item.style.display = match ? '' : 'none';
        });
      });

      if (searchClose) {
        searchClose.addEventListener('click', (e) => {
          e.stopPropagation();
          searchBar.classList.remove('visible');
          searchInput.value = '';
          accordion.querySelectorAll('.document-accordion-item').forEach(item => {
            item.style.display = '';
          });
        });
      }
    }

    // Toggle main accordion on header click
    header.addEventListener('click', (e) => {
      if (e.target.closest('.documents-accordion-search-btn')) return;
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
        if (e.target.closest('.document-copy-btn')) return;
        if (e.target.closest('.document-group-checkbox')) return;

        e.stopPropagation();
        item.classList.toggle('expanded');
        tasksList.style.display = item.classList.contains('expanded') ? 'block' : 'none';
      });

      // Copy button click - copy document link to clipboard
      const copyBtn = item.querySelector('.document-copy-btn');
      if (copyBtn) {
        copyBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          const fileUrl = copyBtn.dataset.fileUrl || '';
          if (fileUrl) {
            navigator.clipboard.writeText(fileUrl);
          }
          // Visual feedback
          const originalContent = copyBtn.innerHTML;
          copyBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#22c55e" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>';
          setTimeout(() => { copyBtn.innerHTML = originalContent; }, 1500);
        });
      }

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
            <button class="task-group-copy-btn" data-file-url="${fileUrl}" title="Copy document link">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                <rect x="9" y="9" width="13" height="13" rx="2" ry="2"></rect>
                <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"></path>
              </svg>
            </button>
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

      const _actionEyeId = extractSlackChannelId(taskMessageLink);
      const _actionEyeHtml = _actionEyeId ? `<span class="todo-slack-eye-btn" data-channel-id="${escapeHtml(_actionEyeId)}" data-channel-name="${escapeHtml(task.participant_text || '')}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;">#</span>` : '';
      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                ${getTrendingIconHtml(task)}${_actionEyeHtml}
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
                    ${getTypeIconHtml(task.type || null, task.participant_text || '') ? `<span style="flex:1;min-width:4px;"></span>${getTypeIconHtml(task.type || null, task.participant_text || '')}` : ''}
                  </div>
                  ${tagsHtml}
                </div>
                ${taskMessageLink && taskMessageLink.includes('mail.google.com') ? `<span class="todo-block-email-btn" data-todo-id="${task.id}" title="Block email participants"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
                ${_actionEyeId ? `<span class="slack-bell-channel-btn" data-channel-id="${escapeHtml(_actionEyeId)}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-left:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
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
        if (e.target.closest('.task-group-copy-btn')) return;
        if (e.target.closest('.task-group-checkbox')) return;

        group.classList.toggle('expanded');
        tasksContainer.style.display = group.classList.contains('expanded') ? 'block' : 'none';
        chevron.style.transform = group.classList.contains('expanded') ? 'rotate(180deg)' : '';
      });

      // Copy button click - copy document link to clipboard
      const groupCopyBtn = group.querySelector('.task-group-copy-btn');
      if (groupCopyBtn) {
        groupCopyBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          const fileUrl = groupCopyBtn.dataset.fileUrl || '';
          if (fileUrl) {
            navigator.clipboard.writeText(fileUrl);
          }
          // Visual feedback
          const originalContent = groupCopyBtn.innerHTML;
          groupCopyBtn.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#22c55e" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>';
          setTimeout(() => { groupCopyBtn.innerHTML = originalContent; }, 1500);
        });
      }

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

        // Eye btn — open channel transcript slider
        const eyeBtn = taskItem.querySelector('.todo-slack-eye-btn');
        if (eyeBtn) {
          eyeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const channelId = eyeBtn.dataset.channelId;
            const chName = eyeBtn.dataset.channelName;
            if (channelId) openChannelTranscriptSlider(channelId, chName);
          });
        }

        // Block email btn
        const blockBtnCh = taskItem.querySelector('.todo-block-email-btn');
        if (blockBtnCh) {
          blockBtnCh.addEventListener('click', (e) => { e.stopPropagation(); showBlockEmailModal(blockBtnCh.dataset.todoId); });
        }

        // Block Slack channel btn
        const blockChBtnTag = taskItem.querySelector('.slack-bell-channel-btn');
        if (blockChBtnTag && !blockChBtnTag.dataset.blockHandlerAttached) {
          blockChBtnTag.dataset.blockHandlerAttached = 'true';
          blockChBtnTag.addEventListener('click', (e) => { e.stopPropagation(); showBellDropdown(e, blockChBtnTag.dataset.channelId); });
        }

        // Task item click (open transcript)
        taskItem.addEventListener('click', (e) => {
          if (e.target.closest('.todo-checkbox')) return;
          if (e.target.closest('.todo-source')) return;
          if (e.target.closest('.todo-slack-eye-btn')) return;
          if (e.target.closest('.todo-block-email-btn')) return;
          if (e.target.closest('.slack-bell-channel-btn')) return;
          if (e.target.closest('a')) return;
          if (e.metaKey || e.ctrlKey || e.shiftKey) {
            e.preventDefault();
            isMultiSelectMode = true;
            toggleTodoSelection(todoId);
            return;
          }
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
    const channelIdForView = extractSlackChannelId(channelLink);

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
            ${channelIdForView ? `<span class="slack-channel-view-btn" data-channel-id="${channelIdForView}" data-channel-name="${escapeHtml(channelName)}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;margin-right:6px;">#</span>` : ""}
            ${channelIdForView ? `<span class="slack-bell-channel-btn" data-channel-id="${channelIdForView}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-right:6px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ""}
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
                ${getTrendingIconHtml(task)}
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
                    ${getTypeIconHtml(task.type || null, task.participant_text || '') ? `<span style="flex:1;min-width:4px;"></span>${getTypeIconHtml(task.type || null, task.participant_text || '')}` : ''}
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


      const _tagEyeId = extractSlackChannelId(taskMessageLink);
      const _tagEyeHtml = _tagEyeId ? `<span class="todo-slack-eye-btn" data-channel-id="${escapeHtml(_tagEyeId)}" data-channel-name="${escapeHtml(task.participant_text || '')}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;">#</span>` : '';
      return `
              <div class="task-group-task-item todo-item ${isUnread ? 'unread' : ''} ${isNew ? 'appearing' : ''}" data-todo-id="${task.id}" data-message-link="${escapeHtml(taskMessageLink)}">
                ${getTrendingIconHtml(task)}${_tagEyeHtml}
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
                    <div class="todo-sources" style="display: flex; align-items: center; gap: 2px; flex: 1; min-width: 0;">
                      ${sourceIconHtml}
                      ${secondaryLinksHtml}
                    </div>
                    ${(getTypeIconHtml(task.type || null, task.participant_text || '') || _tagEyeId) ? `<span style="flex:1;min-width:4px;display:inline-flex;"></span>` : ''}
                    ${getTypeIconHtml(task.type || null, task.participant_text || '') || ''}
                    ${_tagEyeId ? `<span class="slack-bell-channel-btn" data-channel-id="${escapeHtml(_tagEyeId)}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-left:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
                  </div>
                  ${tagsHtml}
                </div>
                ${taskMessageLink && taskMessageLink.includes('mail.google.com') ? `<span class="todo-block-email-btn" data-todo-id="${task.id}" title="Block email participants"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
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
        // Don't toggle if clicking the open button, checkbox, or view button
        if (e.target.closest('.task-group-open-btn')) return;
        if (e.target.closest('.task-group-checkbox')) return;
        if (e.target.closest('.slack-channel-view-btn')) return;
        if (e.target.closest('.slack-bell-channel-btn')) return;

        group.classList.toggle('expanded');
        tasksContainer.style.display = group.classList.contains('expanded') ? 'block' : 'none';
        if (chevron) {
          chevron.style.transform = group.classList.contains('expanded') ? 'rotate(180deg)' : '';
        }
      });

      // Eye/View icon click - open channel-level transcript slider
      const viewBtn = group.querySelector('.slack-channel-view-btn');
      if (viewBtn) {
        viewBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          const channelId = viewBtn.dataset.channelId;
          const chName = viewBtn.dataset.channelName;
          if (channelId) {
            openChannelTranscriptSlider(channelId, chName);
          }
        });
      }

      // Block channel btn - mute Slack channel
      const blockChBtn = group.querySelector('.task-group-header .slack-bell-channel-btn');
      if (blockChBtn && !blockChBtn.dataset.blockHandlerAttached) {
        blockChBtn.dataset.blockHandlerAttached = 'true';
        blockChBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          showBellDropdown(e, blockChBtn.dataset.channelId);
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
          console.log('Slack task item clicked, todoId:', todoId);
          if (e.target.closest('.todo-checkbox')) return;
          if (e.target.closest('.todo-source')) return;
          if (e.target.closest('.todo-tag')) return;
          if (e.target.closest('.todo-slack-eye-btn')) return;
          if (e.target.closest('.todo-block-email-btn')) return;
          if (e.target.closest('.slack-bell-channel-btn')) return;
          if (e.target.closest('.view-more-inline')) return;
          if (e.metaKey || e.ctrlKey || e.shiftKey) {
            e.preventDefault();
            isMultiSelectMode = true;
            toggleTodoSelection(todoId);
            return;
          }
          showTranscriptSlider(todoId);
        });

        // Eye btn — open channel transcript slider
        const eyeBtn = taskItem.querySelector('.todo-slack-eye-btn');
        if (eyeBtn) {
          eyeBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const channelId = eyeBtn.dataset.channelId;
            const chName = eyeBtn.dataset.channelName;
            if (channelId) openChannelTranscriptSlider(channelId, chName);
          });
        }

        // Block email btn
        const blockBtnSl = taskItem.querySelector('.todo-block-email-btn');
        if (blockBtnSl) {
          blockBtnSl.addEventListener('click', (e) => { e.stopPropagation(); showBlockEmailModal(blockBtnSl.dataset.todoId); });
        }

        // Block Slack channel btn in channel group task
        const blockChBtnSl = taskItem.querySelector('.slack-bell-channel-btn');
        if (blockChBtnSl && !blockChBtnSl.dataset.blockHandlerAttached) {
          blockChBtnSl.dataset.blockHandlerAttached = 'true';
          blockChBtnSl.addEventListener('click', (e) => { e.stopPropagation(); showBellDropdown(e, blockChBtnSl.dataset.channelId); });
        }

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
      const fyiChannelId = extractSlackChannelId(sortedTasks[0]?.message_link || '');

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
            ${fyiChannelId ? `<span class="slack-channel-view-btn" data-channel-id="${fyiChannelId}" data-channel-name="${escapeHtml(channelName)}" title="View entire channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;font-size:13px;font-weight:700;color:#667eea;opacity:0.7;cursor:pointer;flex-shrink:0;margin-left:auto;margin-right:4px;">#</span>` : ''}
            ${fyiChannelId ? `<span class="slack-bell-channel-btn" data-channel-id="${fyiChannelId}" title="Mute this channel" style="display:inline-flex;align-items:center;justify-content:center;width:18px;height:18px;cursor:pointer;flex-shrink:0;margin-right:4px;"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg></span>` : ''}
          </div>
          <div class="slack-channel-tasks-list" style="display: none;">
            ${sortedTasks.map(task => {
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
                ${getTrendingIconHtml(task)}
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
    const unreadChannelCount = filteredGroups.filter(group => group.tasks.some(t => isTaskUnread(t.id))).length;
    const displayCount = unreadChannelCount > 0 ? unreadChannelCount : totalChannels;
    const hasAnyUnread = unreadChannelCount > 0;

    return `
      <div class="slack-channels-accordion ${hasAnyUnread ? 'has-unread' : ''}">
        <div class="slack-channels-accordion-header">
          <div class="slack-channels-accordion-title">
            <img src="${slackIconUrl}" alt="Slack" style="width: 16px; height: 16px;">
            <span>Slack Channels</span>
            <span class="slack-channels-accordion-count">${displayCount}</span>
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

  // Update Slack Channels accordion count to reflect channels with unread tasks
  function updateSlackChannelsAccordionCount() {
    document.querySelectorAll('.slack-channels-accordion').forEach(accordion => {
      const items = accordion.querySelectorAll('.slack-channel-accordion-item');
      let unreadChannels = 0;
      items.forEach(item => {
        const taskItems = item.querySelectorAll('.slack-channel-task-item');
        const hasUnread = Array.from(taskItems).some(t => {
          const todoId = t.dataset.todoId;
          return todoId && isTaskUnread(todoId);
        });
        if (hasUnread) {
          item.classList.add('unread');
          unreadChannels++;
        } else {
          item.classList.remove('unread');
        }
      });
      const countEl = accordion.querySelector('.slack-channels-accordion-count');
      if (countEl) {
        countEl.textContent = unreadChannels > 0 ? unreadChannels : items.length;
      }
      if (unreadChannels > 0) {
        accordion.classList.add('has-unread');
      } else {
        accordion.classList.remove('has-unread');
      }
    });
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
            // Remove completed tasks from allFyiItems array to prevent reappearing
            const completedIdSet = new Set(taskIds);
            allFyiItems = allFyiItems.filter(t => !completedIdSet.has(String(t.id)) && !completedIdSet.has(t.id));
            allTodos = allTodos.filter(t => !completedIdSet.has(String(t.id)) && !completedIdSet.has(t.id));
            // Update accordion count
            const countEl = accordion.querySelector('.slack-channels-accordion-count');
            const remainingItems = accordion.querySelectorAll('.slack-channel-accordion-item');
            if (remainingItems.length === 0) {
              accordion.remove();
            } else if (countEl) {
              countEl.textContent = remainingItems.length;
            }
            updateTabCounts();
            updateSlackChannelsAccordionCount();
          }, 400);
        });
      }

      itemHeader.addEventListener('click', (e) => {
        // Don't toggle if clicking the open button, checkbox, or view button
        if (e.target.closest('.slack-channel-open-btn')) return;
        if (e.target.closest('.slack-channel-group-checkbox')) return;
        if (e.target.closest('.slack-channel-view-btn')) return;
        if (e.target.closest('.slack-bell-channel-btn')) return;

        e.stopPropagation();
        item.classList.toggle('expanded');
        tasksList.style.display = item.classList.contains('expanded') ? 'block' : 'none';
      });

      // Block channel btn in FYI accordion header
      const blockChBtnFyi = item.querySelector('.slack-channel-item-header .slack-bell-channel-btn');
      if (blockChBtnFyi && !blockChBtnFyi.dataset.blockHandlerAttached) {
        blockChBtnFyi.dataset.blockHandlerAttached = 'true';
        blockChBtnFyi.addEventListener('click', (e) => {
          e.stopPropagation();
          showBellDropdown(e, blockChBtnFyi.dataset.channelId);
        });
      }

      // Eye/View icon click - open channel-level transcript slider
      const viewBtn = item.querySelector('.slack-channel-view-btn');
      if (viewBtn) {
        viewBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          const channelId = viewBtn.dataset.channelId;
          const chName = viewBtn.dataset.channelName;
          if (channelId) {
            openChannelTranscriptSlider(channelId, chName);
          }
        });
      }
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
          // Remove from allFyiItems to prevent reappearing on re-render
          allFyiItems = allFyiItems.filter(t => String(t.id) !== String(todoId));
          allTodos = allTodos.filter(t => String(t.id) !== String(todoId));
          // Update count or remove channel item if empty
          const remainingTasks = channelItem.querySelectorAll('.slack-channel-task-item').length;
          if (remainingTasks === 0) {
            channelItem.remove();
            // Update accordion count
            const remainingItems = accordion.querySelectorAll('.slack-channel-accordion-item');
            if (remainingItems.length === 0) {
              accordion.remove();
            } else {
              const countEl = accordion.querySelector('.slack-channels-accordion-count');
              if (countEl) countEl.textContent = remainingItems.length;
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
          updateSlackChannelsAccordionCount();
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
      const eyeBtn = item.querySelector('.todo-slack-eye-btn');
      const todoId = item.dataset.todoId;

      // Eye button click - open channel transcript slider
      eyeBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        const channelId = eyeBtn.dataset.channelId;
        const chName = eyeBtn.dataset.channelName;
        if (channelId) openChannelTranscriptSlider(channelId, chName);
      });

      // Block email button click - show block modal
      const blockBtn = item.querySelector('.todo-block-email-btn');
      blockBtn?.addEventListener('click', (e) => {
        e.stopPropagation();
        showBlockEmailModal(blockBtn.dataset.todoId);
      });

      // Block Slack channel button click
      const blockChBtn = item.querySelector('.slack-bell-channel-btn');
      if (blockChBtn && !blockChBtn.dataset.blockHandlerAttached) {
        blockChBtn.dataset.blockHandlerAttached = 'true';
        blockChBtn.addEventListener('click', (e) => {
          e.stopPropagation();
          showBellDropdown(e, blockChBtn.dataset.channelId);
        });
      }

      // Make whole item clickable for transcript or multi-select
      item.addEventListener('click', (e) => {
        // Don't trigger if clicking on interactive elements
        if (e.target.closest('.todo-checkbox') ||
          e.target.closest('.todo-star') ||
          e.target.closest('.todo-clock') ||
          e.target.closest('.todo-slack-eye-btn') ||
          e.target.closest('.todo-block-email-btn') ||
          e.target.closest('.slack-bell-channel-btn') ||
          e.target.closest('.todo-source') ||
          e.target.closest('.todo-tag') ||
          e.target.closest('.view-more-inline') ||
          e.target.closest('a')) {
          return;
        }

        // Command/Ctrl/Shift + click on the item itself triggers multi-select
        if (e.metaKey || e.ctrlKey || e.shiftKey) {
          e.preventDefault();
          isMultiSelectMode = true;
          toggleTodoSelection(todoId);
          return;
        }

        showTranscriptSlider(todoId);
      });

      checkbox?.addEventListener('click', (e) => {
        e.stopPropagation();
        // Check if Command/Ctrl/Shift key is pressed for multi-select
        if (e.metaKey || e.ctrlKey || e.shiftKey) {
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

    // Close current slider
    if (closeSliderFn) {
      closeSliderFn();
    }
  }


  // ─────────────────────────────────────────────────────────────────
  // Shared @mention handler — attach to any (replyInput, replySection) pair
  // ─────────────────────────────────────────────────────────────────
  function setupMentionHandler(replyInput, replySection) {
    let mentionDropdown = null;
    let mentionSearchTimeout = null;
    let mentionSelectedIndex = 0;
    let mentionResults = [];
    let currentMentionQuery = '';
    let mentionAbortController = null;
    let mentionInsertCooldown = false;

    const createMentionDropdown = () => {
      if (mentionDropdown) return mentionDropdown;
      mentionDropdown = document.createElement('div');
      mentionDropdown.className = 'mention-dropdown';
      mentionDropdown.style.display = 'none';
      replySection.style.position = 'relative';
      replySection.insertBefore(mentionDropdown, replySection.firstChild);
      return mentionDropdown;
    };

    const showMentionLoading = () => {
      const dropdown = createMentionDropdown();
      dropdown.innerHTML = `
        <div class="mention-dropdown-header">Searching users...</div>
        <div class="mention-loading"><div class="spinner"></div><span>Loading...</span></div>`;
      dropdown.style.display = 'block';
    };

    const showMentionResults = (users) => {
      const dropdown = createMentionDropdown();
      mentionResults = users;
      mentionSelectedIndex = 0;
      while (dropdown.firstChild) dropdown.removeChild(dropdown.firstChild);
      if (users.length === 0) {
        const h = document.createElement('div'); h.className = 'mention-dropdown-header'; h.textContent = 'Users'; dropdown.appendChild(h);
        const e = document.createElement('div'); e.className = 'mention-empty'; e.textContent = `No users found for "${currentMentionQuery}"`; dropdown.appendChild(e);
      } else {
        const h = document.createElement('div'); h.className = 'mention-dropdown-header'; h.textContent = `Users (${users.length})`; dropdown.appendChild(h);
        const listDiv = document.createElement('div'); listDiv.className = 'mention-dropdown-list';
        users.forEach((user, index) => {
          const initials = (user.name || user.email || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
          const item = document.createElement('div');
          item.className = `mention-item ${index === 0 ? 'selected' : ''}`;
          item.dataset.index = index; item.dataset.name = user.name || user.email || ''; item.dataset.email = user.email || ''; item.dataset.slackId = user.slack_id || '';
          item.innerHTML = `<div class="mention-avatar">${initials}</div><div class="mention-info"><div class="mention-name">${escapeHtml(user.name || user.email || 'Unknown')}</div>${user.email ? `<div class="mention-email">${escapeHtml(user.email)}</div>` : ''}</div>`;
          listDiv.appendChild(item);
        });
        dropdown.appendChild(listDiv);
      }
      dropdown.style.display = 'block';
      void dropdown.offsetHeight;
    };

    const hideMentionDropdown = () => {
      if (mentionDropdown) mentionDropdown.style.display = 'none';
      mentionResults = []; mentionSelectedIndex = 0; currentMentionQuery = '';
    };

    const updateMentionSelection = () => {
      if (!mentionDropdown) return;
      mentionDropdown.querySelectorAll('.mention-item').forEach((item, index) => {
        if (index === mentionSelectedIndex) { item.classList.add('selected'); item.scrollIntoView({ block: 'nearest' }); }
        else item.classList.remove('selected');
      });
    };

    const insertMention = (name, email, slackId) => {
      const mentionSpan = document.createElement('span');
      mentionSpan.className = 'mention-tag'; mentionSpan.contentEditable = 'false';
      mentionSpan.dataset.email = email || ''; mentionSpan.dataset.slackId = slackId || '';
      mentionSpan.textContent = `@${name}`;

      const selection = window.getSelection();
      let replaced = false;
      if (selection.rangeCount > 0) {
        const range = selection.getRangeAt(0);
        const containerNode = range.endContainer;
        if (containerNode.nodeType === Node.TEXT_NODE) {
          const textContent = containerNode.textContent;
          const cursorOffset = range.endOffset;
          let atPos = -1;
          for (let i = cursorOffset - 1; i >= 0; i--) {
            if (textContent[i] === '@') { atPos = i; break; }
            if (textContent[i] === ' ' || textContent[i] === '\n') break;
          }
          if (atPos !== -1) {
            const beforeText = textContent.substring(0, atPos);
            const afterText = textContent.substring(cursorOffset);
            const frag = document.createDocumentFragment();
            if (beforeText) frag.appendChild(document.createTextNode(beforeText));
            frag.appendChild(mentionSpan);
            frag.appendChild(document.createTextNode('\u00A0' + afterText));
            containerNode.parentNode.replaceChild(frag, containerNode);
            replaced = true;
          }
        }
      }
      if (!replaced) { replyInput.appendChild(mentionSpan); replyInput.appendChild(document.createTextNode('\u00A0')); }

      const newRange = document.createRange();
      const sel = window.getSelection();
      newRange.selectNodeContents(replyInput); newRange.collapse(false);
      sel.removeAllRanges(); sel.addRange(newRange);

      mentionInsertCooldown = true;
      setTimeout(() => { mentionInsertCooldown = false; }, 500);
      hideMentionDropdown();
      replyInput.focus();
    };

    // Event delegation for dropdown clicks
    createMentionDropdown().addEventListener('mousedown', (e) => {
      const item = e.target.closest('.mention-item');
      if (item) {
        e.preventDefault(); e.stopPropagation();
        insertMention(item.dataset.name, item.dataset.email, item.dataset.slackId || '');
      }
    });

    const searchMentionUsers = async (query) => {
      if (mentionAbortController) { mentionAbortController.abort(); mentionAbortController = null; }
      currentMentionQuery = query;
      showMentionLoading();
      const controller = new AbortController();
      mentionAbortController = controller;
      const timeoutId = setTimeout(() => controller.abort(), 10000);
      const queryWords = query.toLowerCase().split(/\s+/).filter(w => w.length > 0);
      const backendQuery = queryWords.length > 1 ? queryWords[0] : query;
      try {
        const response = await fetch(WEBHOOK_URL, {
          method: 'POST', headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(createAuthenticatedPayload({ action: 'search_user', query: backendQuery, timestamp: new Date().toISOString() })),
          signal: controller.signal
        });
        clearTimeout(timeoutId); mentionAbortController = null;
        if (!response.ok) throw new Error('Search failed');
        const responseText = await response.text();
        let data = [];
        if (responseText && responseText.trim()) { try { data = JSON.parse(responseText); } catch(e) { data = []; } }
        if (typeof data === 'string') { try { data = JSON.parse(data); } catch(e) { data = []; } }
        let users = Array.isArray(data) ? data : (data?.users || data?.results || []);
        users = users.map(u => ({
          name: u['Full Name'] || u.name || u.full_name || '',
          email: u.user_email_ID || u.email || '',
          slack_id: u.user_slack_ID || u.slack_id || ''
        })).filter(u => u.name || u.email);
        // Client-side multi-word filtering
        if (queryWords.length > 1) {
          users = users.filter(u => {
            const haystack = (u.name + ' ' + u.email).toLowerCase();
            return queryWords.every(w => haystack.includes(w));
          });
        }
        currentMentionQuery = query;
        showMentionResults(users);
      } catch (error) {
        clearTimeout(timeoutId); mentionAbortController = null;
        if (error.name !== 'AbortError') console.error('Mention search error:', error);
        if (query === currentMentionQuery) hideMentionDropdown();
      }
    };

    const checkForMention = () => {
      if (mentionInsertCooldown) return;
      const selection = window.getSelection();
      if (!selection.rangeCount) return;
      const range = selection.getRangeAt(0);
      const containerNode = range.endContainer;
      if (containerNode.nodeType !== Node.TEXT_NODE) {
        if (!mentionAbortController) hideMentionDropdown();
        else if (mentionDropdown) mentionDropdown.style.display = 'none';
        return;
      }
      if (containerNode.parentElement?.classList?.contains('mention-tag')) {
        if (!mentionAbortController) hideMentionDropdown();
        else if (mentionDropdown) mentionDropdown.style.display = 'none';
        return;
      }
      const textContent = containerNode.textContent;
      const cursorOffset = range.endOffset;
      let atPos = -1;
      for (let i = cursorOffset - 1; i >= 0; i--) {
        if (textContent[i] === '@') { atPos = i; break; }
        if (textContent[i] === '\n') break;
      }
      if (atPos !== -1) {
        const query = textContent.substring(atPos + 1, cursorOffset);
        const trimmedQuery = query.trimEnd();
        if (trimmedQuery.length >= 2) {
          if (trimmedQuery === currentMentionQuery && mentionDropdown?.style?.display === 'block') return;
          if (mentionSearchTimeout) clearTimeout(mentionSearchTimeout);
          mentionSearchTimeout = setTimeout(() => searchMentionUsers(trimmedQuery), 400);
        } else if (trimmedQuery.length < 2) {
          hideMentionDropdown();
        }
      } else {
        hideMentionDropdown();
      }
    };

    replyInput.addEventListener('input', checkForMention);
    replyInput.addEventListener('keydown', (e) => {
      if (!mentionDropdown || mentionDropdown.style.display === 'none' || mentionResults.length === 0) return;
      switch (e.key) {
        case 'ArrowDown': e.preventDefault(); mentionSelectedIndex = Math.min(mentionSelectedIndex + 1, mentionResults.length - 1); updateMentionSelection(); break;
        case 'ArrowUp':   e.preventDefault(); mentionSelectedIndex = Math.max(mentionSelectedIndex - 1, 0); updateMentionSelection(); break;
        case 'Enter': if (mentionResults.length > 0) { e.preventDefault(); const s = mentionResults[mentionSelectedIndex]; insertMention(s.name || s.email, s.email || '', s.slack_id || ''); } break;
        case 'Escape': e.preventDefault(); hideMentionDropdown(); break;
        case 'Tab':   if (mentionResults.length > 0) { e.preventDefault(); const s = mentionResults[mentionSelectedIndex]; insertMention(s.name || s.email, s.email || '', s.slack_id || ''); } break;
      }
    }, true);

    document.addEventListener('click', (e) => {
      if (mentionDropdown && !mentionDropdown.contains(e.target) && !replyInput.contains(e.target)) hideMentionDropdown();
    });
  }

  // Detect if message_html has over-bolded content from Slack rich_text block conversion.
  // Slack's rich_text blocks sometimes wrap all text in <b> when converted to HTML by n8n,
  // even though only specific portions should be bold. If most of the HTML text is inside
  // <b>/<strong> tags, prefer the plain text + formatMessageContent path instead.
  function isOverBoldedHtml(messageHtml, messagePlain) {
    if (!messageHtml || !messagePlain) return false;
    // If plain text has Slack mrkdwn patterns, prefer plain text rendering
    if (messagePlain.includes('<<@') || /\*\S[^*\n]+\S\*/.test(messagePlain)) return true;
    const t = document.createElement('div');
    t.innerHTML = messageHtml;
    const boldEls = t.querySelectorAll('b, strong');
    if (boldEls.length === 0) return false;
    let boldTextLen = 0;
    boldEls.forEach(el => { boldTextLen += (el.textContent || '').length; });
    const totalTextLen = (t.textContent || '').length;
    // If >40% of text is bold, it's likely over-bolded from Slack rich_text conversion
    return totalTextLen > 0 && (boldTextLen / totalTextLen) > 0.4;
  }

  // Optimistic message rendering — immediately shows the user's message in the transcript
  function appendOptimisticMessage(messagesContainer, messageText, messageHtml, isDarkMode) {
    if (!messagesContainer) return null;
    // Resolve user name: try profile email (shrey.jain → Shrey Jain), then userDirectory match, then 'You'
    let userName = 'You';
    let avatarUrl = '';
    const email = userProfileData?.email_ID || window.Oracle?.state?.userData?.email || '';
    if (email) {
      // Try to find in the transcript's user_directory stored on the slider
      const slider = messagesContainer.closest('.transcript-slider');
      const storedMessages = slider?.transcriptMessages || [];
      // Find a message from the current user by matching email prefix to message_from
      const emailPrefix = email.split('@')[0]?.toLowerCase(); // e.g. "shrey.jain"
      const nameParts = emailPrefix?.split('.')?.map(w => w.charAt(0).toUpperCase() + w.slice(1)) || [];
      if (nameParts.length > 0) userName = nameParts.join(' ');
      // Try to find avatar from transcript messages sent by this user
      const myMsg = storedMessages.find(m => {
        const from = (m.message_from || '').toLowerCase();
        return nameParts.some(p => from.includes(p.toLowerCase()));
      });
      if (myMsg?.avatar_url) avatarUrl = myMsg.avatar_url;
    }
    const chipColor = typeof getUserChipColor === 'function' ? getUserChipColor(userName) : '#667eea';
    const initials = userName.split(/\s+/).map(w => w.charAt(0)).join('').toUpperCase().slice(0, 2);
    const avatarHtml = avatarUrl
      ? `<img src="${avatarUrl}" style="width:32px;height:32px;border-radius:50%;object-fit:cover;flex-shrink:0;" onerror="this.style.display='none';this.nextElementSibling.style.display='flex';" /><div style="width:32px;height:32px;background:${chipColor};border-radius:50%;display:none;align-items:center;justify-content:center;color:white;font-size:11px;font-weight:600;flex-shrink:0;">${initials}</div>`
      : `<div style="width:32px;height:32px;background:${chipColor};border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:${initials.length > 1 ? '11' : '13'}px;font-weight:600;flex-shrink:0;">${initials}</div>`;
    const msgDiv = document.createElement('div');
    msgDiv.className = 'transcript-message optimistic-message';
    msgDiv.style.cssText = 'display:flex;flex-direction:column;gap:6px;opacity:0.7;transition:opacity 0.3s;';
    const displayContent = messageHtml
      ? (typeof sanitizeHtml === 'function' ? sanitizeHtml(messageHtml) : messageHtml)
      : (typeof formatMessageContent === 'function' ? formatMessageContent(messageText) : (typeof escapeHtml === 'function' ? escapeHtml(messageText) : messageText));
    msgDiv.innerHTML = `
      <div style="display:flex;align-items:center;gap:8px;">
        ${avatarHtml}
        <div style="flex:1;min-width:0;">
          <div style="font-weight:600;font-size:13px;color:${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${typeof escapeHtml === 'function' ? escapeHtml(userName) : userName}</div>
          <div style="font-size:11px;color:${isDarkMode ? '#666' : '#95a5a6'};">Just now</div>
        </div>
      </div>
      <div style="margin-left:40px;padding:12px 16px;background:${isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border-radius:12px;border-top-left-radius:4px;font-size:14px;color:${isDarkMode ? '#e8e8e8' : '#2c3e50'};line-height:1.6;word-wrap:break-word;">${displayContent}</div>`;
    messagesContainer.appendChild(msgDiv);
    messagesContainer.scrollTop = messagesContainer.scrollHeight;
    return msgDiv;
  }

  // Fun loading quotes for transcript loading screens
  const loadingQuotes = [
    "Honey never spoils — 3,000-year-old honey from Egyptian tombs is still edible 🍯",
    "Octopuses have three hearts and blue blood 🐙",
    "Russia spans 11 time zones — more than any other country 🇷🇺",
    "The shortest war in history lasted 38 minutes — Britain vs Zanzibar, 1896 ⚔️",
    "Venus is the only planet that spins clockwise 🪐",
    "There are more trees on Earth than stars in the Milky Way 🌳",
    "Cleopatra lived closer to the Moon landing than to the Great Pyramid 🏛️",
    "A day on Venus is longer than a year on Venus 🌅",
    "The Amazon River has no bridges crossing it 🌊",
    "Oxford University is older than the Aztec Empire 🎓",
    "Bananas are berries, but strawberries aren't 🍌",
    "The Sahara Desert was green and lush just 6,000 years ago 🌿",
    "Iceland has no mosquitoes 🦟❌",
    "The Great Wall of China isn't visible from space, but Spanish greenhouses are 🛰️",
    "Sharks predate trees by 50 million years 🦈",
    "Finland has more saunas than cars 🧖",
    "The Dead Sea is so salty you float without trying 🏊",
    "A group of flamingos is called a 'flamboyance' 🦩",
    "Lake Baikal holds 20% of the world's unfrozen freshwater 💧",
    "The Eiffel Tower grows ~6 inches taller in summer due to heat expansion 🗼",
    "Greenland is the largest island but has the lowest population density 🏔️",
    "Lightning strikes the Earth about 100 times every second ⚡",
    "Wombat poop is cube-shaped 🟫",
    "The Library of Alexandria held an estimated 400,000 scrolls 📜",
    "Antarctica is technically a desert — the driest continent on Earth 🏜️",
    "Bhutan measures Gross National Happiness instead of GDP 😊",
    "A jiffy is an actual unit of time — 1/100th of a second ⏱️",
    "There are more possible chess games than atoms in the observable universe ♟️",
    "The human nose can detect over 1 trillion different scents 👃",
    "Scotland's national animal is the unicorn 🦄",
    "There are more public libraries in the US than McDonald's restaurants 📚",
    "The inventor of the Pringles can is buried in one 🥔",
    "A cloud can weigh more than a million pounds ☁️",
    "The longest hiccuping spree lasted 68 years 😮",
    "Cows have best friends and get stressed when separated 🐄",
    "The total weight of all ants on Earth roughly equals all humans 🐜",
    "Hot water freezes faster than cold water — the Mpemba effect 🧊",
    "Your brain uses 20% of your body's total energy ⚡🧠",
    "The Mariana Trench is deeper than Everest is tall 🌊",
    "Polar bears' fur is transparent, not white 🐻‍❄️",
    "A teaspoon of neutron star weighs about 6 billion tons ⭐",
    "Dolphins sleep with one eye open 🐬",
    "The world's oldest known recipe is for beer — from 4,000 years ago 🍺",
    "Your body has more bacterial cells than human cells 🦠",
    "The oldest known living tree is over 5,000 years old 🌲",
    "Sound travels 4.3 times faster in water than in air 🔊",
    "There's a basketball court on the top floor of the US Supreme Court 🏀",
    "The unicorn is Scotland's national animal since the 12th century 🏴󠁧󠁢󠁳󠁣󠁴󠁿",
    "Pluto hasn't completed a full orbit since its discovery in 1930 🪐",
    "A bolt of lightning is five times hotter than the sun's surface ☀️",
    "Humans share 60% of their DNA with bananas 🧬",
    "The Great Barrier Reef is the largest living structure on Earth 🪸",
    "Rome is older than Italy — by about 2,600 years 🇮🇹",
    "There are more stars in the universe than grains of sand on Earth 🌌",
    "The longest place name has 85 letters — it's a hill in New Zealand 🇳🇿",
    "Sloths can hold their breath longer than dolphins — up to 40 minutes 🦥",
    "The shortest commercial flight lasts 57 seconds — in Scotland ✈️",
    "Only 5% of the ocean has been explored 🌊",
    "A day on Mercury lasts 59 Earth days ☿️",
    "The tallest mountain in the solar system is on Mars — Olympus Mons 🏔️",
    "Honey bees can recognize human faces 🐝",
    "There are more than 7,000 languages spoken worldwide today 🗣️",
    "The speed of light could circle the Earth 7.5 times per second 💡",
    "Saudi Arabia imports camels from Australia 🐪",
    "A group of owls is called a 'parliament' 🦉",
    "The longest river in Asia is the Yangtze at 6,300 km 🏞️",
    "Astronauts grow up to 2 inches taller in space 🧑‍🚀",
    "The human eye can distinguish about 10 million different colors 👁️",
    "Papua New Guinea has over 840 languages — the most of any country 🗺️",
    "Mammoths were still alive when the Great Pyramid was built 🦣",
    "An average person walks about 100,000 miles in a lifetime 🚶",
    "The longest English word without a vowel is 'rhythms' 🔤",
    "The Pacific Ocean is larger than all landmasses combined 🌏",
    "Butterflies taste with their feet 🦋",
    "The human brain can store roughly 2.5 petabytes of data 💾",
    "Canada has more lakes than the rest of the world combined 🇨🇦",
    "A jellyfish is 95% water 🪼",
    "The world's quietest room is in Minneapolis — you can hear your blood flow 🤫",
    "Mount Everest grows about 4mm taller every year 📐",
    "The first email was sent in 1971 — the message was 'QWERTYUIOP' 📧",
    "Sea otters hold hands while sleeping to avoid drifting apart 🦦",
    "There are 118 ridges on the edge of a US dime 🪙",
    "The Mona Lisa has no eyebrows — it was the fashion in Renaissance Florence 🎨",
    "Earth's core is as hot as the surface of the Sun — about 5,500°C 🌍",
    "A single strand of spider silk is stronger than a steel wire of the same width 🕷️",
    "The average person spends 6 months of their life waiting at red lights 🚦",
    "Koalas sleep up to 22 hours a day 🐨",
    "The longest wedding veil was longer than 63 football fields 👰",
    "More people have been to the Moon than to the bottom of the Mariana Trench 🌙",
    "A flea can jump 150 times its own body length 🦗",
    "The inventor of the fire hydrant is unknown — the patent was lost in a fire 🚒",
    "Your stomach gets a new lining every 3 to 4 days 🫁",
    "The world's largest snowflake was 15 inches wide — found in Montana, 1887 ❄️",
    "Trees can communicate with each other through underground fungal networks 🍄",
    "The longest game of Monopoly lasted 70 straight days 🎲",
    "A cockroach can live for a week without its head 🪳",
    "The first oranges weren't orange — they were green 🍊",
    "Astronauts' footprints on the Moon will last millions of years 🌕",
    "The Netflix DVD library once held over 100,000 titles 📀",
    "Cats have over 20 vocalizations, including the purr, which no one fully understands 🐱",
    "The total length of your blood vessels could wrap around Earth twice 🫀",
  ];
  function getLoadingQuote() { return loadingQuotes[Math.floor(Math.random() * loadingQuotes.length)]; }
  function startRotatingQuote(el) {
    if (!el) return;
    let idx = Math.floor(Math.random() * loadingQuotes.length);
    el.textContent = loadingQuotes[idx];
    const timer = setInterval(() => {
      if (!el.isConnected) { clearInterval(timer); return; }
      idx = (idx + 1) % loadingQuotes.length;
      el.style.opacity = '0';
      setTimeout(() => { el.textContent = loadingQuotes[idx]; el.style.opacity = '1'; }, 100);
    }, 3000);
    el._quoteTimer = timer;
  }

  // Quick Reply Slider — opens immediately from in-memory message data, no network call.
  // Shows a single message and lets the user reply to it as a thread.
  function openReplySlider(channelId, channelName, msg, isDarkMode) {
    isTranscriptSliderOpen = true;
    document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

    const col3Rect = window.Oracle.getCol3Rect();
    const slackIconUrl = chrome.runtime.getURL('icon-slack.png');
    const threadTs = msg.ts || msg.time;

    const overlay = document.createElement('div');
    overlay.className = 'transcript-slider-overlay';
    if (col3Rect) {
      overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out; transition: width 0.3s ease, left 0.3s ease;`;
    } else {
      overlay.style.cssText = 'position: fixed; top: 0; right: 0; bottom: 0; width: 400px; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;';
    }

    const slider = document.createElement('div');
    slider.className = 'transcript-slider';
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; overflow-x: hidden; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 16px 20px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0;`;
    header.innerHTML = `
      <div style="display: flex; align-items: center; gap: 12px;">
        <div style="width: 36px; height: 36px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 10px; display: flex; align-items: center; justify-content: center;"><img src="${slackIconUrl}" alt="Slack" style="width: 20px; height: 20px;"></div>
        <div>
          <div style="font-weight: 600; font-size: 15px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Reply to Thread</div>
          <div style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">${escapeHtml(channelName || '')}</div>
        </div>
      </div>
      <div style="display: flex; gap: 8px; align-items: center;">
        <button class="transcript-expand-btn" title="Expand" style="background: none; border: none; cursor: pointer; color: ${isDarkMode ? '#888' : '#7f8c8d'}; font-size: 18px; padding: 4px 8px;">⛶</button>
        <button class="transcript-close-btn" title="Close" style="background: none; border: none; cursor: pointer; color: ${isDarkMode ? '#888' : '#7f8c8d'}; font-size: 20px; padding: 4px 8px;">×</button>
      </div>`;
    slider.appendChild(header);

    // Message container
    const msgContainer = document.createElement('div');
    msgContainer.style.cssText = `flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; gap: 8px;`;

    // ── Render using the same pipeline as the full transcript slider ──
    const msgDiv = document.createElement('div');
    msgDiv.className = 'transcript-message';
    msgDiv.style.cssText = 'display: flex; flex-direction: column; gap: 6px;';

    const displayTime = formatTimeAgoFresh(msg.time || msg.ts || '');
    const chipColor = getUserChipColor(msg.user_id || msg.message_from || 'Unknown');
    const nameParts = (msg.message_from || 'U').trim().split(/\s+/);
    const initials = nameParts.length > 1
      ? (nameParts[0].charAt(0) + nameParts[nameParts.length - 1].charAt(0)).toUpperCase()
      : nameParts[0].charAt(0).toUpperCase();
    const avatarUrl = msg.avatar_url;
    const initialsFallback = `<div style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`;
    const avatarHtml = avatarUrl
      ? `<img src="${escapeHtml(avatarUrl)}" class="transcript-avatar-img" style="width: 32px; height: 32px; border-radius: 50%; object-fit: cover; flex-shrink: 0;" /><div class="transcript-avatar-fallback" style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: none; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`
      : initialsFallback;

    const msgHeader = document.createElement('div');
    msgHeader.style.cssText = 'display: flex; align-items: center; gap: 8px;';
    msgHeader.innerHTML = `
      ${avatarHtml}
      <div style="flex: 1; min-width: 0;">
        <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
        <div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${displayTime}</div>
      </div>`;

    // Avatar fallback handler
    const avatarImg = msgHeader.querySelector('.transcript-avatar-img');
    if (avatarImg) {
      avatarImg.addEventListener('error', () => {
        avatarImg.style.display = 'none';
        const fb = avatarImg.nextElementSibling;
        if (fb) fb.style.display = 'flex';
      });
    }

    const msgContent = document.createElement('div');
    msgContent.className = 'transcript-message-content';
    msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-weight: 400; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word;`;
    // Use message_html if available (preserves hyperlinks from emails), fallback to plain text
    // Skip message_html if it has over-bolded content from Slack rich_text conversion
    const htmlContent = (msg.message_html && !isOverBoldedHtml(msg.message_html, msg.message)) ? msg.message_html : '';
    if (htmlContent) {
      const isSimple = (() => {
        const t = document.createElement('div');
        t.innerHTML = htmlContent;
        const text = t.textContent || '';
        if (text.length > 500) return false;
        if (t.querySelectorAll('table, img, iframe, object, embed').length > 0) return false;
        if ((htmlContent.match(/style="/g) || []).length > 5) return false;
        if (t.querySelectorAll('hr, blockquote').length > 0) return false;
        return true;
      })();
      if (isSimple) {
        // Simple HTML with links — sanitize and render inline
        msgContent.innerHTML = typeof sanitizeHtml === 'function' ? sanitizeHtml(htmlContent) : htmlContent;
      } else {
        msgContent.dataset.emailIframe = 'true';
        msgContent.dataset.rawHtml = htmlContent;
      }
    } else {
      msgContent.innerHTML = formatMessageContent(msg.message || '');
    }

    msgDiv.appendChild(msgHeader);
    msgDiv.appendChild(msgContent);

    // Attachments
    if (msg.attachments && msg.attachments.length > 0) {
      const validAttachments = msg.attachments.filter(att => {
        if (att.type === 'link' && (!att.name || att.name === 'Attachment')) return false;
        if (!att.name && !att.url && !att.text) return false;
        return true;
      });
      if (validAttachments.length > 0) {
        const attachmentsDiv = document.createElement('div');
        attachmentsDiv.style.cssText = 'margin-left: 40px; margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px;';
        validAttachments.forEach(att => {
          const el = renderTranscriptAttachment(att);
          if (el) attachmentsDiv.appendChild(el);
        });
        if (attachmentsDiv.children.length > 0) msgDiv.appendChild(attachmentsDiv);
      }
    }

    // Reactions
    if (msg.reactions && msg.reactions.length > 0) {
      const reactionsDiv = document.createElement('div');
      reactionsDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px;';
      msg.reactions.forEach(reaction => {
        const emojiName = reaction.emoji || reaction.name || '';
        const count = reaction.count || 1;
        const users = reaction.users || [];
        const emojiUnicode = convertSlackEmoji(emojiName);
        const tooltipText = users.length > 0 ? users.map(u => typeof u === 'object' ? (u.name || u.user_id) : u).join(', ') : `${count} reaction${count > 1 ? 's' : ''}`;
        const pill = document.createElement('span');
        pill.style.cssText = `display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: ${isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'}; border: 1px solid ${isDarkMode ? 'rgba(102,126,234,0.25)' : 'rgba(102,126,234,0.15)'}; color: ${isDarkMode ? '#b0b0b0' : '#555'}; cursor: default; position: relative;`;
        pill.innerHTML = `<span style="font-size: 14px;">${emojiUnicode}</span>${count > 1 ? `<span style="font-size: 11px; font-weight: 500;">${count}</span>` : ''}`;
        const tooltip = document.createElement('div');
        tooltip.style.cssText = `display: none; position: absolute; bottom: calc(100% + 6px); left: 50%; transform: translateX(-50%); padding: 6px 10px; border-radius: 8px; font-size: 11px; max-width: 280px; white-space: normal; z-index: 10001; pointer-events: none; background: ${isDarkMode ? '#2a2a2a' : '#333'}; color: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.25);`;
        tooltip.innerHTML = formatReactionTooltip(tooltipText);
        pill.appendChild(tooltip);
        pill.addEventListener('mouseenter', () => tooltip.style.display = 'block');
        pill.addEventListener('mouseleave', () => tooltip.style.display = 'none');
        reactionsDiv.appendChild(pill);
      });
      msgDiv.appendChild(reactionsDiv);
    }

    // Thread indicator (existing replies)
    if (msg.thread && msg.thread.reply_count > 0) {
      const replyCount = msg.thread.reply_count;
      const replyUsers = msg.thread.reply_users || [];
      const latestReply = msg.thread.latest_reply || '';
      let avatarsHtml = '';
      replyUsers.slice(0, 3).forEach((user, i) => {
        const userName = typeof user === 'object' ? (user.name || 'U') : (typeof user === 'string' ? user : 'U');
        const initial = userName.charAt(0).toUpperCase();
        const uColor = getUserChipColor(typeof user === 'object' ? (user.user_id || userName) : userName);
        avatarsHtml += `<div style="width: 20px; height: 20px; border-radius: 50%; background: ${uColor}; display: flex; align-items: center; justify-content: center; color: white; font-size: 9px; font-weight: 600; margin-left: ${i > 0 ? '-4px' : '0'}; border: 1.5px solid ${isDarkMode ? '#1f2940' : 'white'}; position: relative; z-index: ${3 - i};">${initial}</div>`;
      });
      const latestReplyText = latestReply ? ` · Last reply ${formatTimeAgoFresh(latestReply)}` : '';
      const threadDiv = document.createElement('div');
      threadDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; align-items: center; gap: 6px;';
      threadDiv.innerHTML = `<div style="display: flex; align-items: center;">${avatarsHtml}</div><span style="font-size: 12px; color: #667eea; font-weight: 600;">${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}</span><span style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${latestReplyText}</span>`;
      msgDiv.appendChild(threadDiv);
    }

    msgContainer.appendChild(msgDiv);
    slider.appendChild(msgContainer);

    // Reply section — matches style of the full transcript slider
    const replySection = document.createElement('div');
    replySection.style.cssText = `padding: 16px 20px; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; background: ${isDarkMode ? 'rgba(26,26,46,0.8)' : 'rgba(248,249,250,0.8)'}; flex-shrink: 0;`;
    replySection.innerHTML = `
      <div style="display: flex; gap: 12px; align-items: flex-end;">
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button class="channel-mic-btn" title="Voice to text" style="background: rgba(102,126,234,0.1); border: 1px solid rgba(102,126,234,0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">🎤</button>
          <input type="file" class="channel-file-input" multiple style="display: none;">
          <button class="channel-attach-btn" title="Attach files" style="background: rgba(102,126,234,0.1); border: 1px solid rgba(102,126,234,0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">📎</button>
        </div>
        <div class="transcript-reply-input" contenteditable="true" placeholder="Reply to thread..." style="flex: 1; padding: 12px 16px; border: 2px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 12px; font-family: inherit; font-size: 14px; min-height: 44px; max-height: 120px; overflow-y: auto; outline: none; background: ${isDarkMode ? '#16213e' : 'white'}; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.5;"></div>
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button type="button" class="channel-oracle-btn" title="Ask Oracle Assistant" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 34px; height: 34px; border-radius: 10px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.2s; padding: 0; font-size: 14px;">∞</button>
          <button class="reply-send-btn" title="Send reply (⌘+Enter)" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 34px; height: 34px; border-radius: 12px; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; flex-shrink: 0;"><span style="font-size: 18px;">➤</span></button>
        </div>
      </div>
      <div style="margin-top: 8px; font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">Press Enter to type, ⌘+Enter to send</div>`;
    slider.appendChild(replySection);

    // Mic, attach, oracle buttons wired up
    const micBtn = replySection.querySelector('.channel-mic-btn');
    const attachBtn = replySection.querySelector('.channel-attach-btn');
    const fileInput = replySection.querySelector('.channel-file-input');
    const oracleBtn = replySection.querySelector('.channel-oracle-btn');
    if (micBtn) micBtn.addEventListener('click', () => { const ri = replySection.querySelector('.transcript-reply-input'); if (typeof startVoiceRecording === 'function') startVoiceRecording(ri, micBtn); });
    if (attachBtn && fileInput) attachBtn.addEventListener('click', () => fileInput.click());
    if (oracleBtn) oracleBtn.addEventListener('click', () => { if (typeof window.OracleAssistant?.showChatSlider === 'function') window.OracleAssistant.showChatSlider({ mode: 'fullscreen' }); });

    overlay.appendChild(slider);
    document.body.appendChild(overlay);

    // Close
    const closeSlider = () => {
      isTranscriptSliderOpen = false;
      slider.style.animation = 'slideOutRight 0.3s ease-out';
      setTimeout(() => { overlay.remove(); window.Oracle.collapseCol3AfterSlider(); }, 300);
    };
    header.querySelector('.transcript-close-btn').addEventListener('click', closeSlider);
    document.addEventListener('keydown', function escHandler(e) {
      if (e.key === 'Escape') { if (document.querySelector('.attachment-preview-modal')) return; e.stopImmediatePropagation(); closeSlider(); document.removeEventListener('keydown', escHandler); }
    });
    header.querySelector('.transcript-expand-btn').addEventListener('click', () => {
      window.Oracle.expandCol3ForSlider();
      const newRect = window.Oracle.getCol3Rect();
      if (newRect) { overlay.style.left = newRect.left + 'px'; overlay.style.width = newRect.width + 'px'; }
    });

    // Send reply
    const sendBtn = replySection.querySelector('.reply-send-btn');
    const replyInput = replySection.querySelector('.transcript-reply-input');
    if (typeof setupRichTextFormatting === 'function') setupRichTextFormatting(replyInput);
    setupMentionHandler(replyInput, replySection);

    const sendReply = async () => {
      const replyText = replyInput?.innerText?.trim() || '';
      const replyTextSlack = (typeof convertContentEditableToSlackMrkdwn === 'function')
        ? convertContentEditableToSlackMrkdwn(replyInput)
        : replyText;
      if (!replyTextSlack) return;
      // Optimistic: show message immediately
      const replyHtmlContent = replyInput.innerHTML;
      appendOptimisticMessage(msgContainer, replyText, replyHtmlContent, isDarkMode);
      replyInput.innerHTML = '';
      sendBtn.disabled = true;
      sendBtn.innerHTML = '<span style="font-size: 14px;">⏳</span>';
      try {
        const payload = createAuthenticatedPayload({
          action: 'reply_to_thread',
          channel_id: channelId,
          thread_ts: threadTs,
          reply_text: replyTextSlack,
          reply_text_plain: replyText,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-newtab'
        });
        const response = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
        if (response.ok) {
          sendBtn.style.background = 'linear-gradient(45deg, #27ae60, #2ecc71)';
          sendBtn.innerHTML = '<span style="font-size: 14px;">✓</span>';
          // Make optimistic message fully opaque
          const optMsg = msgContainer.querySelector('.optimistic-message:last-child');
          if (optMsg) optMsg.style.opacity = '1';
          // Reset send button after delay (no slider replacement — optimistic message stays visible)
          setTimeout(() => {
            sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
            sendBtn.innerHTML = '<span style="font-size: 18px;">➤</span>';
            sendBtn.disabled = false;
          }, 1500);
        } else { throw new Error('Send failed'); }
      } catch (err) {
        console.error('Reply error:', err);
        // Remove optimistic message on failure
        const optMsg = msgContainer.querySelector('.optimistic-message:last-child');
        if (optMsg) { optMsg.style.opacity = '0'; setTimeout(() => optMsg.remove(), 300); }
        replyInput.innerHTML = replyHtmlContent;
        sendBtn.style.background = 'rgba(231,76,60,0.2)';
        sendBtn.innerHTML = '<span style="font-size: 14px;">✗</span>';
        setTimeout(() => { sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)'; sendBtn.innerHTML = '<span style="font-size: 18px;">➤</span>'; sendBtn.disabled = false; }, 2000);
      }
    };

    sendBtn.addEventListener('click', sendReply);
    replyInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); sendReply(); }
    });

    // Expand col3 and focus input
    window.Oracle.expandCol3ForSlider();
    setTimeout(() => replyInput.focus(), 100);
  }

  // Channel-level transcript slider (no thread, just channel_id)
  // Opens a transcript slider for the entire channel/DM/group DM
  // =============================================
  async function openChannelTranscriptSlider(channelId, channelName, threadTs = null, threadMessageLink = null) {
    console.log('Opening channel transcript slider for:', channelId, channelName);

    isTranscriptSliderOpen = true;

    // Remove any existing transcript slider
    document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
    document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());

    const col3Rect = window.Oracle.getCol3Rect();
    const isDarkMode = document.body.classList.contains('dark-mode');
    let isExpanded = false;
    const slackIconUrl = chrome.runtime.getURL('icon-slack.png');

    const overlay = document.createElement('div');
    overlay.className = 'transcript-slider-overlay';
    if (col3Rect) {
      overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out; transition: width 0.3s ease, left 0.3s ease;`;
    } else {
      overlay.style.cssText = 'position: fixed; top: 0; right: 0; bottom: 0; width: 400px; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;';
    }

    const slider = document.createElement('div');
    slider.className = 'transcript-slider';
    slider.style.cssText = `width: 100%; height: 100%; background: ${isDarkMode ? '#1f2940' : 'white'}; box-shadow: -4px 0 20px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); display: flex; flex-direction: column; animation: slideInRight 0.3s ease-out; transition: width 0.3s ease; overflow-x: hidden; border-radius: 12px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};`;

    // Header
    const header = document.createElement('div');
    header.style.cssText = `padding: 20px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; display: flex; align-items: center; justify-content: space-between; flex-shrink: 0;`;
    header.innerHTML = `
      <div style="display: flex; align-items: center; gap: 12px;">
        <div style="width: 40px; height: 40px; background: linear-gradient(45deg, #667eea, #764ba2); border-radius: 10px; display: flex; align-items: center; justify-content: center;"><img src="${slackIconUrl}" alt="Slack" style="width: 24px; height: 24px;"></div>
        <div style="flex: 1; min-width: 0;">
          <div style="font-weight: 600; font-size: 16px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${threadTs ? 'Thread' : escapeHtml(channelName || 'Channel')}</div>
          <div class="transcript-message-count" style="font-size: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};"><span class="channel-name-sub">${threadTs ? escapeHtml(channelName || '') : ''}</span>${threadTs ? ' · ' : ''}Loading ${threadTs ? 'thread' : 'channel'} messages...</div>
        </div>
      </div>
      <div style="display: flex; align-items: center; gap: 8px;">
        <button class="transcript-expand-btn" title="Expand" style="background: ${isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)'}; border: none; color: #667eea; width: 36px; height: 36px; border-radius: 8px; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">⬅</button>
        <button class="transcript-close-btn" style="background: rgba(231,76,60,0.1); border: none; color: #e74c3c; width: 36px; height: 36px; border-radius: 8px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">×</button>
      </div>
    `;
    slider.appendChild(header);

    // Messages container with loading spinner
    const messagesContainer = document.createElement('div');
    messagesContainer.className = 'transcript-messages-container';
    messagesContainer.style.cssText = 'flex: 1; overflow-y: auto; overflow-x: hidden; padding: 20px; display: flex; flex-direction: column; gap: 16px;';
    messagesContainer.innerHTML = `
      <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 16px;">
        <div class="spinner" style="width: 40px; height: 40px; border: 4px solid rgba(102, 126, 234, 0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
        <div class="loading-quote" style="font-size: 14px; color: #7f8c8d; transition: opacity 0.1s;">${getLoadingQuote()}</div>
      </div>
    `;
    slider.appendChild(messagesContainer);
    startRotatingQuote(messagesContainer.querySelector('.loading-quote'));
    const replySection = document.createElement('div');
    replySection.style.cssText = `padding: 16px 20px; border-top: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'}; background: ${isDarkMode ? 'rgba(26, 26, 46, 0.8)' : 'rgba(248,249,250,0.8)'}; flex-shrink: 0;`;
    replySection.innerHTML = `
      <div class="channel-attachments" style="display: none; margin-bottom: 10px; padding: 10px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.1)' : 'rgba(102, 126, 234, 0.05)'}; border-radius: 8px; border: 1px dashed rgba(102, 126, 234, 0.3);">
        <div class="channel-attachments-list" style="display: flex; flex-wrap: wrap; gap: 8px;"></div>
      </div>
      <div style="display: flex; gap: 12px; align-items: flex-end;">
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button class="channel-mic-btn" title="Voice to text" style="background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">🎤</button>
          <input type="file" class="channel-file-input" multiple style="display: none;">
          <button class="channel-attach-btn" title="Attach files" style="background: rgba(102, 126, 234, 0.1); border: 1px solid rgba(102, 126, 234, 0.3); color: #667eea; width: 44px; height: 38px; border-radius: 10px; cursor: pointer; font-size: 18px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">📎</button>
        </div>
        <div class="transcript-reply-input" contenteditable="true" placeholder="${threadTs ? 'Reply to thread...' : 'Type a message to the channel...'}" style="flex: 1; padding: 12px 16px; border: 2px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 12px; font-family: inherit; font-size: 14px; min-height: 44px; max-height: 120px; overflow-y: auto; transition: all 0.2s; outline: none; background: ${isDarkMode ? '#16213e' : 'white'}; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.5;"></div>
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button type="button" class="channel-oracle-btn" title="Ask Oracle Assistant" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 34px; height: 34px; border-radius: 10px; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.2s; padding: 0; font-size: 14px;">∞</button>
          <button class="channel-reply-send-btn" title="${threadTs ? 'Send reply (Cmd+Enter)' : 'Send to channel (Cmd+Enter)'}" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 34px; height: 34px; border-radius: 12px; cursor: pointer; font-size: 16px; display: flex; align-items: center; justify-content: center; transition: all 0.2s; flex-shrink: 0;"><span style="font-size: 18px;">➤</span></button>
        </div>
      </div>
      <div style="margin-top: 8px; font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">Press Enter to type, ⌘+Enter to send</div>
    `;
    slider.appendChild(replySection);

    // --- Attachment handling for channel slider ---
    const channelAttachBtn = slider.querySelector('.channel-attach-btn');
    const channelFileInput = slider.querySelector('.channel-file-input');
    const channelAttachmentsContainer = slider.querySelector('.channel-attachments');
    const channelAttachmentsList = slider.querySelector('.channel-attachments-list');
    let channelPendingAttachments = [];

    if (channelAttachBtn && channelFileInput) {
      channelAttachBtn.addEventListener('click', () => channelFileInput.click());
      channelFileInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        for (const file of files) {
          if (file.size > 10 * 1024 * 1024) { alert(`File "${file.name}" exceeds 10MB limit.`); continue; }
          const base64 = await fileToBase64(file);
          channelPendingAttachments.push({ name: file.name, type: file.type, size: file.size, data: base64 });
          const attachmentEl = document.createElement('div');
          attachmentEl.style.cssText = `display: flex; align-items: center; gap: 6px; background: ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'white'}; padding: 6px 10px; border-radius: 6px; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225, 232, 237, 0.6)'}; font-size: 12px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};`;
          const icon = getFileIcon(file.type);
          attachmentEl.innerHTML = `<span>${icon}</span><span style="max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${escapeHtml(file.name)}</span><span style="color: #95a5a6; font-size: 10px;">(${formatFileSize(file.size)})</span><button class="remove-ch-attachment" data-name="${escapeHtml(file.name)}" style="background: none; border: none; color: #e74c3c; cursor: pointer; font-size: 14px; padding: 0 4px;">×</button>`;
          channelAttachmentsList.appendChild(attachmentEl);
          attachmentEl.querySelector('.remove-ch-attachment').addEventListener('click', () => {
            channelPendingAttachments = channelPendingAttachments.filter(a => a.name !== file.name);
            attachmentEl.remove();
            if (channelPendingAttachments.length === 0) channelAttachmentsContainer.style.display = 'none';
          });
        }
        if (channelPendingAttachments.length > 0) channelAttachmentsContainer.style.display = 'block';
        channelFileInput.value = '';
      });
    }

    // --- Mic button for channel slider ---
    const channelMicBtn = slider.querySelector('.channel-mic-btn');
    if (channelMicBtn) {
      channelMicBtn.addEventListener('click', () => {
        const rInput = slider.querySelector('.transcript-reply-input');
        if (typeof startVoiceRecording === 'function') startVoiceRecording(rInput, channelMicBtn);
      });
    }

    // Rich text formatting for channel reply input
    const channelReplyInput = slider.querySelector('.transcript-reply-input');
    if (channelReplyInput) setupRichTextFormatting(channelReplyInput);
    if (channelReplyInput) setupMentionHandler(channelReplyInput, replySection);

    // --- Oracle button for channel slider ---
    const channelOracleBtn = slider.querySelector('.channel-oracle-btn');
    if (channelOracleBtn) {
      channelOracleBtn.addEventListener('click', () => {
        if (typeof window.OracleAssistant?.showChatSlider === 'function') {
          window.OracleAssistant.showChatSlider({ mode: 'fullscreen' });
        }
      });
    }

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
    document.addEventListener('keydown', function escHandler(e) {
      if (e.key === 'Escape') { if (document.querySelector('.attachment-preview-modal')) return; e.stopImmediatePropagation(); closeSlider(); document.removeEventListener('keydown', escHandler); }
    });

    // Expand/collapse handler
    slider.querySelector('.transcript-expand-btn').addEventListener('click', () => {
      isExpanded = !isExpanded;
      if (isExpanded) {
        window.Oracle.expandCol3ForSlider();
      } else {
        window.Oracle.collapseCol3AfterSlider();
      }
      const newRect = window.Oracle.getCol3Rect();
      if (newRect) {
        overlay.style.left = newRect.left + 'px';
        overlay.style.width = newRect.width + 'px';
      }
    });

    // Send reply handler (unthreaded - channel-level message)
    const sendBtn = slider.querySelector('.channel-reply-send-btn');
    const replyInput = slider.querySelector('.transcript-reply-input');

    const sendChannelReply = async () => {
      const replyText = replyInput?.innerText?.trim() || '';
      const replyTextSlack = convertContentEditableToSlackMrkdwn(replyInput);
      if (!replyTextSlack && channelPendingAttachments.length === 0) return;

      // Optimistic: show message immediately
      const replyHtmlContent = replyInput.innerHTML;
      appendOptimisticMessage(messagesContainer, replyText, replyHtmlContent, isDarkMode);
      replyInput.innerHTML = '';
      sendBtn.disabled = true;
      sendBtn.innerHTML = '<span style="font-size: 14px;">⏳</span>';

      try {
        const payload = createAuthenticatedPayload({
          action: threadTs ? 'reply_to_thread' : 'reply_to_channel',
          channel_id: channelId,
          thread_ts: threadTs || undefined,
          message_link: threadMessageLink || undefined,
          reply_text: replyTextSlack,
          reply_text_plain: replyText,
          attachments: channelPendingAttachments.length > 0 ? channelPendingAttachments : undefined,
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
          // Make optimistic message fully opaque
          const optMsg = messagesContainer.querySelector('.optimistic-message:last-child');
          if (optMsg) optMsg.style.opacity = '1';
          channelPendingAttachments = [];
          if (channelAttachmentsList) channelAttachmentsList.innerHTML = '';
          if (channelAttachmentsContainer) channelAttachmentsContainer.style.display = 'none';
          // Reload transcript after short delay to let Slack process the message
          setTimeout(() => {
            sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
            sendBtn.innerHTML = '<span style="font-size: 18px;">➤</span>';
            sendBtn.disabled = false;
            loadChannelMessages();
          }, 1500);
        } else {
          throw new Error('Failed to send message');
        }
      } catch (error) {
        console.error('Error sending channel message:', error);
        // Remove optimistic message on failure
        const optMsg = messagesContainer.querySelector('.optimistic-message:last-child');
        if (optMsg) { optMsg.style.opacity = '0'; setTimeout(() => optMsg.remove(), 300); }
        replyInput.innerHTML = replyHtmlContent;
        sendBtn.style.background = 'rgba(231,76,60,0.2)';
        sendBtn.innerHTML = '<span style="font-size: 14px;">✗</span>';
        setTimeout(() => {
          sendBtn.style.background = 'linear-gradient(45deg, #667eea, #764ba2)';
          sendBtn.innerHTML = '<span style="font-size: 18px;">➤</span>';
          sendBtn.disabled = false;
        }, 2000);
      }
    };

    sendBtn.addEventListener('click', sendChannelReply);
    replyInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) {
        e.preventDefault();
        sendChannelReply();
      }
      // Plain Enter outside a list → insert line break (not <div>)
      // Skip if already handled by setupRichTextFormatting (e.g. empty list item exit)
      if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.shiftKey && !e.defaultPrevented) {
        const sel = window.getSelection();
        const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
        const inList = node && (node.nodeType === Node.TEXT_NODE
          ? node.parentElement?.closest('ol, ul')
          : node.closest?.('ol, ul'));
        if (!inList) {
          e.preventDefault();
          document.execCommand('insertLineBreak');
        }
      }
    });

    // Reusable function to fetch and render channel messages
    const loadChannelMessages = async () => {
      // Show loading state
      messagesContainer.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 16px;">
          <div class="spinner" style="width: 40px; height: 40px; border: 4px solid rgba(102, 126, 234, 0.2); border-top-color: #667eea; border-radius: 50%; animation: spin 0.8s linear infinite;"></div>
          <div class="loading-quote" style="font-size: 14px; color: #7f8c8d; transition: opacity 0.1s;">${getLoadingQuote()}</div>
        </div>
      `;
      startRotatingQuote(messagesContainer.querySelector('.loading-quote'));

      try {
        const response = await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(createAuthenticatedPayload({
            action: threadTs ? 'fetch_task_details' : 'fetch_channel_transcript',
            channel_id: channelId,
            thread_ts: threadTs || undefined,
            message_link: threadMessageLink || undefined,
            timestamp: new Date().toISOString(),
            source: 'oracle-chrome-extension-newtab'
          }))
        });

      if (!response.ok) throw new Error('Failed to fetch channel messages');

      const responseText = await response.text();
      let data = {};
      if (responseText && responseText.trim()) {
        try { data = JSON.parse(responseText); } catch (e) { data = {}; }
      }

      // Parse response — same format as showTranscriptSlider
      const responseData = Array.isArray(data) ? data[0] : data;
      const transcript = responseData?.transcript || [];
      const userDirectory = responseData?.user_directory || {};

      let messages = [];
      if (transcript && transcript.length > 0) {
        try {
          const flatTranscript = transcript.flat();
          messages = flatTranscript.map(msgStr => {
            try {
              return typeof msgStr === 'string' ? JSON.parse(msgStr) : msgStr;
            } catch { return null; }
          }).filter(m => m !== null);
        } catch (e) {
          console.error('Error parsing channel transcript:', e);
        }
      }

      // Update message count
      const countEl = slider.querySelector('.transcript-message-count');
      if (countEl) countEl.textContent = `${messages.length} message${messages.length !== 1 ? 's' : ''}`;

      if (messages.length === 0) {
        messagesContainer.innerHTML = `
          <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">
            <div style="font-size: 40px;">💬</div>
            <div style="font-size: 14px;">No recent messages in this channel</div>
          </div>
        `;
      } else {
        // Render messages — reuse same rendering as showTranscriptSlider
        messagesContainer.innerHTML = '';

        messages.forEach((msg, msgIndex) => {
          const msgDiv = document.createElement('div');
          msgDiv.className = 'transcript-message';
          msgDiv.style.cssText = 'display: flex; flex-direction: column; gap: 6px;';

          const displayTime = formatTimeAgoFresh(msg.time || '');

          // Avatar + name header
          const chipColor = getUserChipColor(msg.user_id || msg.message_from || 'Unknown');
          const avatarUrl = msg.avatar_url || (userDirectory[msg.user_id]?.avatar_url);
          const nameParts = (msg.message_from || 'U').trim().split(/\s+/);
          const initials = nameParts.length > 1
            ? (nameParts[0].charAt(0) + nameParts[nameParts.length - 1].charAt(0)).toUpperCase()
            : nameParts[0].charAt(0).toUpperCase();
          const initialsFallback = `<div style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`;
          const avatarHtml = avatarUrl
            ? `<img src="${escapeHtml(avatarUrl)}" class="transcript-avatar-img" data-fallback-initials="${initials}" data-fallback-color="${chipColor}" data-fallback-size="${initials.length > 1 ? '11' : '13'}" style="width: 32px; height: 32px; border-radius: 50%; object-fit: cover; flex-shrink: 0;" /><div class="transcript-avatar-fallback" style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: none; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`
            : initialsFallback;

          const msgHeader = document.createElement('div');
          msgHeader.style.cssText = 'display: flex; align-items: center; gap: 8px;';
          msgHeader.innerHTML = `
            ${avatarHtml}
            <div style="flex: 1; min-width: 0;">
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
              <div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${displayTime}</div>
            </div>
          `;

          const msgContent = document.createElement('div');
          msgContent.className = 'transcript-message-content';
          msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-weight: 400; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word;`;
          // Use message_html if available (preserves hyperlinks from emails)
          // Skip message_html if it has over-bolded content from Slack rich_text conversion
          const htmlContent2 = (msg.message_html && !isOverBoldedHtml(msg.message_html, msg.message)) ? msg.message_html : '';
          if (htmlContent2) {
            const isSimple2 = (() => {
              const t = document.createElement('div');
              t.innerHTML = htmlContent2;
              const text = t.textContent || '';
              if (text.length > 500) return false;
              if (t.querySelectorAll('table, img, iframe, object, embed').length > 0) return false;
              if ((htmlContent2.match(/style="/g) || []).length > 5) return false;
              if (t.querySelectorAll('hr, blockquote').length > 0) return false;
              return true;
            })();
            if (isSimple2) {
              msgContent.innerHTML = typeof sanitizeHtml === 'function' ? sanitizeHtml(htmlContent2) : htmlContent2;
            } else {
              msgContent.dataset.emailIframe = 'true';
              msgContent.dataset.rawHtml = htmlContent2;
            }
          } else {
            msgContent.innerHTML = formatMessageContent(msg.message || '');
          }

          msgDiv.appendChild(msgHeader);
          msgDiv.appendChild(msgContent);

          // Render attachments if present
          if (msg.attachments && msg.attachments.length > 0) {
            const validAttachments = msg.attachments.filter(att => {
              if (att.type === 'link' && (!att.name || att.name === 'Attachment')) return false;
              if (!att.name && !att.url && !att.text) return false;
              return true;
            });
            if (validAttachments.length > 0) {
              const attachmentsDiv = document.createElement('div');
              attachmentsDiv.className = 'transcript-message-attachments';
              attachmentsDiv.style.cssText = 'margin-left: 40px; margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px;';
              validAttachments.forEach(attachment => {
                const attachEl = renderTranscriptAttachment(attachment);
                if (attachEl) attachmentsDiv.appendChild(attachEl);
              });
              if (attachmentsDiv.children.length > 0) msgDiv.appendChild(attachmentsDiv);
            }
          }

          // Render reactions if present
          if (msg.reactions && msg.reactions.length > 0) {
            const reactionsDiv = document.createElement('div');
            reactionsDiv.className = 'transcript-message-reactions';
            reactionsDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px;';
            msg.reactions.forEach(reaction => {
              const emojiName = reaction.emoji || reaction.name || '';
              const count = reaction.count || 1;
              const users = reaction.users || [];
              const emojiUnicode = convertSlackEmoji(emojiName);
              const tooltipText = users.length > 0 ? users.map(u => typeof u === 'object' ? (u.name || u.user_id) : u).join(', ') : `${count} reaction${count > 1 ? 's' : ''}`;
              const pill = document.createElement('span');
              pill.className = 'reaction-pill';
              pill.style.cssText = `display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border: 1px solid ${isDarkMode ? 'rgba(102, 126, 234, 0.25)' : 'rgba(102, 126, 234, 0.15)'}; color: ${isDarkMode ? '#b0b0b0' : '#555'}; cursor: default; position: relative;`;
              pill.innerHTML = `<span style="font-size: 14px;">${emojiUnicode}</span>${count > 1 ? `<span style="font-size: 11px; font-weight: 500;">${count}</span>` : ''}`;
              const tooltip = document.createElement('div');
              tooltip.style.cssText = `display: none; position: absolute; bottom: calc(100% + 6px); left: 50%; transform: translateX(-50%); padding: 6px 10px; border-radius: 8px; font-size: 11px; line-height: 1.4; max-width: 280px; white-space: normal; z-index: 10001; pointer-events: none; background: ${isDarkMode ? '#2a2a2a' : '#333'}; color: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.25);`;
              tooltip.innerHTML = formatReactionTooltip(tooltipText);
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
            threadDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; align-items: center; gap: 6px; cursor: pointer;';
            const replyUsers = msg.thread.reply_users || [];
            const replyCount = msg.thread.reply_count;
            const latestReply = msg.thread.latest_reply || '';
            let avatarsHtml = '';
            replyUsers.slice(0, 3).forEach((user, i) => {
              const userName = typeof user === 'object' ? (user.name || 'U') : (typeof user === 'string' ? user : 'U');
              const initial = userName.charAt(0).toUpperCase();
              const userChipColor = getUserChipColor(typeof user === 'object' ? (user.user_id || userName) : userName);
              avatarsHtml += `<div style="width: 20px; height: 20px; border-radius: 50%; background: ${userChipColor}; display: flex; align-items: center; justify-content: center; color: white; font-size: 9px; font-weight: 600; margin-left: ${i > 0 ? '-4px' : '0'}; border: 1.5px solid ${isDarkMode ? '#1f2940' : 'white'}; position: relative; z-index: ${3 - i};">${initial}</div>`;
            });
            const latestReplyText = latestReply ? ` · Last reply ${formatTimeAgoFresh(latestReply)}` : '';
            threadDiv.innerHTML = `
              <div style="display: flex; align-items: center;">${avatarsHtml}</div>
              <span class="thread-replies-link" style="font-size: 12px; color: #667eea; font-weight: 600; cursor: pointer; text-decoration: underline dotted;">${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}</span>
              <span style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${latestReplyText}</span>
            `;
            threadDiv.addEventListener('click', () => openThreadPanel(slider, channelId, channelName, msg.ts || msg.time, msg.message_from || 'Unknown', msg.message || '', isDarkMode));
            msgDiv.appendChild(threadDiv);
          }

          // "Reply" label - shown on hover for non-thread messages only
          if (!(msg.thread && msg.thread.reply_count > 0)) {
            const replyIconBtn = document.createElement('div');
            replyIconBtn.className = 'msg-reply-btn';
            replyIconBtn.title = 'Reply in thread';
            replyIconBtn.textContent = 'Reply';
            replyIconBtn.style.cssText = `position: absolute; bottom: 6px; right: 8px; font-size: 11px; font-weight: 600; color: #667eea; opacity: 0; cursor: pointer; transition: opacity 0.1s; z-index: 2; background: ${isDarkMode ? '#1f2940' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(102,126,234,0.3)' : 'rgba(102,126,234,0.25)'}; border-radius: 6px; padding: 2px 8px; line-height: 1.6;`;
            replyIconBtn.addEventListener('click', (e) => {
              e.stopPropagation();
              document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
              document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());
              isTranscriptSliderOpen = false;
              openReplySlider(channelId, channelName, msg, isDarkMode);
            });
            msgDiv.style.position = 'relative';
            msgDiv.appendChild(replyIconBtn);
            msgDiv.addEventListener('mouseenter', () => replyIconBtn.style.opacity = '1');
            msgDiv.addEventListener('mouseleave', () => replyIconBtn.style.opacity = '0');
          }

          messagesContainer.appendChild(msgDiv);
        });

        // CSP-safe avatar image fallback
        messagesContainer.querySelectorAll('img.transcript-avatar-img').forEach(img => {
          img.addEventListener('error', () => {
            img.style.display = 'none';
            const fallback = img.nextElementSibling;
            if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
              fallback.style.display = 'flex';
            }
          });
          img.addEventListener('load', () => {
            const fallback = img.nextElementSibling;
            if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
              fallback.style.display = 'none';
            }
          });
        });

        // Scroll to bottom
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
      }
    } catch (error) {
      console.error('Error fetching channel messages:', error);
      messagesContainer.innerHTML = `
        <div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100%; gap: 12px; color: ${isDarkMode ? '#888' : '#7f8c8d'};">
          <div style="font-size: 40px;">⚠️</div>
          <div style="font-size: 14px;">Failed to load channel messages</div>
          <div style="font-size: 12px;">Please try again later</div>
        </div>
      `;
    }
    };

    // Initial load
    loadChannelMessages();
  }

  // =============================================
  // Thread panel — delegates to openChannelTranscriptSlider with threadTs
  // Reuses the full transcript UI: message chips, avatars, attachments, Oracle, mic, etc.
  // =============================================
  function openThreadPanel(parentSlider, channelId, channelName, threadTs, authorName, parentText, isDarkMode) {
    if (!channelId || !threadTs) return;
    const slackDomain = getSlackWorkspaceDomain();
    const buildLink = (chId, ts) => {
      if (!chId || !ts) return '';
      const floatMatch = String(ts).match(/^(\d{10})\.(\d{1,6})$/);
      if (floatMatch) {
        return `https://${slackDomain}/archives/${chId}/p${floatMatch[1]}${floatMatch[2].padEnd(6,'0')}`;
      }
      return '';
    };
    const threadMessageLink = buildLink(channelId, threadTs);
    openChannelTranscriptSlider(channelId, channelName, threadTs, threadMessageLink);
  }

  async function showTranscriptSlider(todoId) {
    // Find the todo with this ID (check allTodos, allFyiItems, allCalendarItems, and allCompletedTasks)
    const todo = allTodos.find(t => t.id == todoId) || allFyiItems.find(t => t.id == todoId) || allCalendarItems.find(t => t.id == todoId) || allCompletedTasks.find(t => t.id == todoId);
    if (!todo) return;

    // Mark task as read when slider opens
    markTaskAsRead(todoId);

    // Highlight the active task row with an orange outline
    document.querySelectorAll('.todo-item.slider-active').forEach(el => {
      el.classList.remove('slider-active');
      el.style.outline = '';
      el.style.boxShadow = '';
      el.style.background = '';
    });
    document.querySelectorAll(`[data-todo-id="${todoId}"]`).forEach(el => {
      if (!el.classList.contains('todo-item')) return;
      el.classList.add('slider-active');
      el.style.boxShadow = 'inset 0 0 0 2px #f39c12, 0 0 0 0px transparent';
      el.style.background = 'rgba(243, 156, 18, 0.06)';
      el.style.borderRadius = '8px';
    });

    // Update Slack channels accordion count to reflect read state
    updateSlackChannelsAccordionCount();

    // Also update the visual read state on the task item itself
    document.querySelectorAll(`[data-todo-id="${todoId}"]`).forEach(el => {
      el.classList.remove('unread');
    });

    // Set flag to prevent auto-refresh
    isTranscriptSliderOpen = true;

    // Remove any existing transcript slider (but NOT chat sliders or discussion sliders)
    document.querySelectorAll('.transcript-slider-overlay:not(.discussion-slider-overlay)').forEach(s => s.remove());
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
        <div style="font-weight: 600; font-weight: 400; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(todo.task_title)}</div>
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
        <div class="loading-quote" style="font-size: 14px; color: #7f8c8d; transition: opacity 0.1s;">${getLoadingQuote()}</div>
      </div>
    `;
    slider.appendChild(messagesContainer);
    startRotatingQuote(messagesContainer.querySelector('.loading-quote'));

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
            <div style="width: 24px; height: 24px; border-radius: 6px; display: flex; align-items: center; justify-content: center; overflow: hidden;"><img src="${chrome.runtime.getURL('icon-oracle.png')}" style="width:24px;height:24px;border-radius:6px;"></div>
            <span style="font-weight: 600; font-weight: 400; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</span>
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
      <div style="display: flex; gap: 8px; align-items: flex-end;">
        <div class="transcript-reply-input" contenteditable="true" placeholder="Type your reply..." style="flex: 1; padding: 12px 16px; border: 2px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'}; border-radius: 12px; font-family: inherit; font-size: 14px; min-height: 80px; max-height: 50vh; overflow-y: auto; transition: all 0.2s; outline: none; background: ${isDarkMode ? '#16213e' : 'white'}; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.5;"></div>
        <div style="display: flex; flex-direction: column; gap: 6px; align-items: center;">
          <button class="transcript-send-btn" title="Send (⌘+Enter)" style="background: linear-gradient(45deg, #667eea, #764ba2); color: white; border: none; width: 34px; height: 34px; border-radius: 10px; font-size: 14px; font-weight: 600; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; justify-content: center;">
            ↗
          </button>
          <button class="transcript-done-btn" title="Mark as Done & Close" style="background: rgba(39,174,96,0.12); border: 1px solid rgba(39,174,96,0.3); color: #27ae60; width: 34px; height: 34px; border-radius: 10px; font-size: 16px; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; justify-content: center;">
            ✓
          </button>
        </div>
      </div>
      <div style="display: flex; align-items: center; justify-content: space-between; margin-top: 8px;">
        <div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">Press Enter to type, ⌘+Enter to send</div>
        <div style="display: flex; gap: 4px; align-items: center;">
          <button class="transcript-mic-btn" title="Voice to text" style="background: rgba(102, 126, 234, 0.08); border: 1px solid rgba(102, 126, 234, 0.2); color: #667eea; width: 30px; height: 30px; border-radius: 8px; cursor: pointer; font-size: 14px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">
            🎤
          </button>
          <button class="transcript-recipients-btn" title="Manage Recipients" style="background: rgba(102, 126, 234, 0.08); border: 1px solid rgba(102, 126, 234, 0.2); color: #667eea; width: 30px; height: 30px; border-radius: 8px; cursor: pointer; font-size: 13px; display: none; align-items: center; justify-content: center; transition: all 0.2s;">
            👤
          </button>
          <input type="file" class="transcript-file-input" multiple style="display: none;">
          <button class="transcript-attach-btn" title="Attach files" style="background: rgba(102, 126, 234, 0.08); border: 1px solid rgba(102, 126, 234, 0.2); color: #667eea; width: 30px; height: 30px; border-radius: 8px; cursor: pointer; font-size: 14px; display: flex; align-items: center; justify-content: center; transition: all 0.2s;">
            📎
          </button>
          <button type="button" class="transcript-oracle-btn" title="Ask Oracle Assistant" style="background: linear-gradient(45deg, #667eea, #764ba2); border: none; color: white; width: 30px; height: 30px; border-radius: 50%; cursor: pointer; display: flex; align-items: center; justify-content: center; transition: all 0.2s; padding: 0; overflow: hidden;">
            <img src="${chrome.runtime.getURL('icon-oracle.png')}" style="width:30px;height:30px;border-radius:50%;object-fit:cover;pointer-events:none;">
          </button>
        </div>
      </div>
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
      // Remove active task highlight
      document.querySelectorAll('.todo-item.slider-active').forEach(el => {
        el.classList.remove('slider-active');
        el.style.outline = '';
        el.style.boxShadow = '';
        el.style.background = '';
      });
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

      // Command+Enter behavior: always send the reply if content exists
      if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') {
        e.preventDefault();
        e.stopPropagation(); // Prevent global Cmd+Enter from marking tasks as done

        // Send the reply if there's content (regardless of focus state)
        const replyContent = replyInput?.textContent?.trim() || '';
        if (replyContent) {
          const sendBtn = slider.querySelector('.transcript-send-btn');
          if (sendBtn && !sendBtn.disabled) {
            sendBtn.click();
          }
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
        e.stopImmediatePropagation();
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
      // Extract user directory for consistent name chip rendering (from Slack transcript)
      const userDirectory = responseData?.user_directory || {};
      console.log('Participants extracted:', participants, 'User directory:', Object.keys(userDirectory).length);

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

      // Extract and display related tasks
      const relatedTasks = responseData?.related_tasks || { count: 0, task_ids: [], tasks: [] };
      if (relatedTasks.count > 0) {
        // Add related tasks badge next to the source icon in header
        const headerBtns = slider.querySelector('.transcript-expand-btn')?.parentElement;
        if (headerBtns) {
          const relatedBtn = document.createElement('button');
          relatedBtn.className = 'related-tasks-btn';
          relatedBtn.title = `${relatedTasks.count} related task${relatedTasks.count !== 1 ? 's' : ''}`;
          relatedBtn.style.cssText = `background: ${isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)'}; border: none; color: #667eea; height: 36px; border-radius: 8px; cursor: pointer; font-size: 12px; font-weight: 600; display: flex; align-items: center; justify-content: center; gap: 4px; transition: all 0.2s; padding: 0 10px;`;
          relatedBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg><span>${relatedTasks.count}</span>`;
          // Insert before expand btn
          headerBtns.insertBefore(relatedBtn, headerBtns.firstChild);

          relatedBtn.addEventListener('click', () => {
            // Toggle related tasks overlay
            let existingOverlay = slider.querySelector('.related-tasks-overlay');
            if (existingOverlay) {
              existingOverlay.remove();
              return;
            }

            const rtOverlay = document.createElement('div');
            rtOverlay.className = 'related-tasks-overlay';
            rtOverlay.style.cssText = `position: absolute; top: 70px; right: 12px; width: 320px; max-height: 400px; overflow-y: auto; background: ${isDarkMode ? '#1a1f2e' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(200,210,220,0.8)'}; border-radius: 12px; box-shadow: 0 8px 24px rgba(0,0,0,${isDarkMode ? '0.4' : '0.15'}); z-index: 200; animation: fadeIn 0.15s ease-out;`;

            let listHtml = `<div style="padding: 12px 16px; border-bottom: 1px solid ${isDarkMode ? 'rgba(255,255,255,0.08)' : 'rgba(200,210,220,0.5)'}; display: flex; align-items: center; justify-content: space-between;">
              <span style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${relatedTasks.count} Related Task${relatedTasks.count !== 1 ? 's' : ''}</span>
              <button class="related-tasks-close" style="background: none; border: none; color: ${isDarkMode ? '#888' : '#7f8c8d'}; cursor: pointer; font-size: 16px; padding: 2px 6px;">×</button>
            </div><div style="padding: 8px;">`;

            relatedTasks.tasks.forEach(rt => {
              const iconData = (window.OracleIcons.getIconForLink || (() => ({icon:'🔗'})))(rt.message_link);
              const participantHtml = rt.participant_text ? `<div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'}; margin-top: 2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escapeHtml(rt.participant_text)}</div>` : '';
              listHtml += `<div class="related-task-item" data-task-id="${rt.id}" style="padding: 10px 12px; border-radius: 8px; cursor: pointer; display: flex; align-items: center; gap: 10px; transition: background 0.15s; margin-bottom: 2px;">
                <div style="width: 28px; height: 28px; border-radius: 6px; background: ${isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'}; display: flex; align-items: center; justify-content: center; flex-shrink: 0;">${iconData.icon}</div>
                <div style="flex: 1; min-width: 0;">
                  <div style="font-size: 13px; font-weight: 500; color: ${isDarkMode ? '#e0e0e0' : '#2c3e50'}; line-height: 1.4; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;">${escapeHtml(rt.task_title)}</div>
                  ${participantHtml}
                </div>
              </div>`;
            });

            listHtml += '</div>';
            rtOverlay.innerHTML = listHtml;
            slider.style.position = 'relative';
            slider.appendChild(rtOverlay);

            // Close button
            rtOverlay.querySelector('.related-tasks-close').addEventListener('click', (e) => {
              e.stopPropagation();
              rtOverlay.remove();
            });

            // Hover effect and click on items
            rtOverlay.querySelectorAll('.related-task-item').forEach(item => {
              item.addEventListener('mouseenter', () => {
                item.style.background = isDarkMode ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.06)';
              });
              item.addEventListener('mouseleave', () => {
                item.style.background = 'transparent';
              });
              // Click to open related task's transcript
              item.addEventListener('click', (e) => {
                e.stopPropagation();
                e.preventDefault();
                const taskId = item.dataset.taskId;
                console.log('🔗 Opening related task:', taskId);
                rtOverlay.remove();
                
                // The related task may not exist in allTodos/allFyiItems — inject it temporarily
                const rt = relatedTasks.tasks.find(t => String(t.id) === String(taskId));
                if (rt) {
                  const alreadyExists = allTodos.find(t => t.id == taskId) || allFyiItems.find(t => t.id == taskId) || (typeof allCompletedTasks !== 'undefined' && allCompletedTasks.find(t => t.id == taskId));
                  if (!alreadyExists) {
                    // Create a minimal todo object so showTranscriptSlider can find it
                    allFyiItems.push({
                      id: parseInt(taskId),
                      task_title: rt.task_title || '',
                      task_name: rt.task_title || '',
                      message_link: rt.message_link || '',
                      participant_text: rt.participant_text || '',
                      status: 0,
                      starred: 0,
                      _isRelatedTask: true
                    });
                  }
                }
                
                setTimeout(() => {
                  showTranscriptSlider(taskId);
                }, 150);
              });
            });

            // Close on click outside
            const closeOnOutside = (e) => {
              // If overlay was already removed (by item click), just clean up listener
              if (!document.contains(rtOverlay)) {
                document.removeEventListener('click', closeOnOutside);
                return;
              }
              if (!rtOverlay.contains(e.target) && !relatedBtn.contains(e.target)) {
                rtOverlay.remove();
                document.removeEventListener('click', closeOnOutside);
              }
            };
            setTimeout(() => document.addEventListener('click', closeOnOutside), 100);
          });
        }
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

        // Extract source message timestamp from Slack message_link for highlighting
        // e.g. https://fwbuzz.slack.com/archives/C0AF6T1NL9Y/p1771992431162819
        // The 'p' identifier encodes a Unix timestamp: p1771992431162819 → 1771992431.162819
        let sourceMessageTs = null;
        const slackLink = todo.message_link || '';
        const slackTsMatch = slackLink.match(/\/p(\d{10})(\d{6})/);
        if (slackTsMatch) {
          sourceMessageTs = parseFloat(slackTsMatch[1] + '.' + slackTsMatch[2]);
        }

        // Helper: parse msg.time string (e.g. "25 Feb 2026, 9:36:58 am") to Unix timestamp
        function msgTimeToUnix(timeStr) {
          if (!timeStr) return null;
          const trimmed = timeStr.trim();
          // ISO format
          if (trimmed.match(/^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}/)) {
            const d = trimmed.includes('Z') || trimmed.match(/[+-]\d{2}:\d{2}$/) ? new Date(trimmed) : new Date(trimmed + 'Z');
            return isNaN(d.getTime()) ? null : d.getTime() / 1000;
          }
          // "DD Mon YYYY, HH:MM:SS am/pm" format
          const match = trimmed.match(/(\d{1,2})\s+(\w{3,})\s+(\d{4}),?\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?/i);
          if (match) {
            const [, day, month, year, hour, min, sec, ampm] = match;
            const months = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
            let h = parseInt(hour);
            if (ampm) { if (ampm.toLowerCase() === 'pm' && h !== 12) h += 12; else if (ampm.toLowerCase() === 'am' && h === 12) h = 0; }
            const mi = months[month.toLowerCase().substring(0, 3)];
            if (mi !== undefined) {
              // Assume IST (UTC+5:30) since timestamps are in Indian time
              const d = new Date(Date.UTC(parseInt(year), mi, parseInt(day), h - 5, parseInt(min) - 30, parseInt(sec || 0)));
              return isNaN(d.getTime()) ? null : d.getTime() / 1000;
            }
          }
          return null;
        }

        // Determine which message index matches the source message link
        // Only highlight if it's NOT the last message (last message is already visible)
        let sourceMessageIndex = -1;
        slider._sourceMessageIndex = -1;
        if (sourceMessageTs) {
          messages.forEach((msg, idx) => {
            const msgTs = msgTimeToUnix(msg.time);
            if (msgTs && Math.abs(msgTs - sourceMessageTs) < 2) { // 2-second tolerance
              sourceMessageIndex = idx;
              slider._sourceMessageIndex = idx;
            }
          });
        }
        const shouldHighlightSource = sourceMessageIndex >= 0;

        messages.forEach((msg, msgIndex) => {
          const isLastMessage = msgIndex === messages.length - 1;
          const isCollapsed = isEmailTask && !isLastMessage && messages.length > 1 && !(shouldHighlightSource && msgIndex === sourceMessageIndex);
          const isSourceMessage = shouldHighlightSource && msgIndex === sourceMessageIndex;

          const msgDiv = document.createElement('div');
          msgDiv.className = `transcript-message ${isCollapsed ? 'collapsed' : ''}${isSourceMessage ? ' source-message-highlight' : ''}`;
          msgDiv.style.cssText = `display: flex; flex-direction: column; gap: 6px;${isSourceMessage ? ` background: ${isDarkMode ? 'rgba(255, 214, 0, 0.1)' : 'rgba(255, 214, 0, 0.15)'}; border-left: 3px solid #ffd600; border-radius: 8px; padding: 8px; margin-left: -8px; margin-right: -8px; position: relative;` : ''}`;

          // Add "Source" badge to highlighted message
          if (isSourceMessage) {
            const badge = document.createElement('div');
            badge.style.cssText = `position: absolute; top: 4px; right: 8px; font-size: 10px; font-weight: 600; color: ${isDarkMode ? '#ffd600' : '#b8960c'}; background: ${isDarkMode ? 'rgba(255, 214, 0, 0.15)' : 'rgba(255, 214, 0, 0.25)'}; padding: 1px 6px; border-radius: 4px; letter-spacing: 0.5px; text-transform: uppercase;`;
            badge.textContent = 'Source';
            msgDiv.appendChild(badge);
          }

          // Always use the actual message timestamp from the conversation thread.
          // Previously used todo.updated_at for the last message, but that reflects
          // sync/task-update time, not when the message was actually sent — causing
          // misleading "Just now" for old messages across Slack, Gmail, and Drive.
          let displayTime = formatTimeAgoFresh(msg.time || '');

          const msgHeader = document.createElement('div');
          msgHeader.style.cssText = `display: flex; align-items: center; gap: 8px; ${isCollapsed ? 'cursor: pointer;' : ''}`;

          // Generate consistent color from user_id for name chips
          const chipColor = getUserChipColor(msg.user_id || msg.message_from || 'Unknown');
          const avatarUrl = msg.avatar_url || (userDirectory[msg.user_id]?.avatar_url);
          // Generate initials: first letter of first name + first letter of last name
          const nameParts = (msg.message_from || 'U').trim().split(/\s+/);
          const initials = nameParts.length > 1
            ? (nameParts[0].charAt(0) + nameParts[nameParts.length - 1].charAt(0)).toUpperCase()
            : nameParts[0].charAt(0).toUpperCase();
          const initialsFallback = `<div style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`;
          const avatarHtml = avatarUrl
            ? `<img src="${escapeHtml(avatarUrl)}" class="transcript-avatar-img" data-fallback-initials="${initials}" data-fallback-color="${chipColor}" data-fallback-size="${initials.length > 1 ? '11' : '13'}" style="width: 32px; height: 32px; border-radius: 50%; object-fit: cover; flex-shrink: 0;" /><div class="transcript-avatar-fallback" style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: none; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`
            : initialsFallback;

          msgHeader.innerHTML = `
            ${isCollapsed ? '<div class="collapse-indicator" style="color: #667eea; font-size: 10px; margin-right: -4px;">▶</div>' : ''}
            ${avatarHtml}
            <div style="flex: 1; min-width: 0;">
              <div style="font-weight: 600; font-size: 13px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
              <div style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${displayTime}</div>
            </div>
          `;

          const msgContent = document.createElement('div');
          msgContent.className = 'transcript-message-content';
          // Updated styling to support HTML tables and rich content
          msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-weight: 400; font-size: 14px; color: ${isDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; overflow-x: hidden; overflow-y: visible; max-width: 100%; ${isCollapsed ? 'display: none;' : ''}`;

          // Format the message content (preserves HTML for emails, formats plain text for Slack)
          // Skip message_html if it has over-bolded content from Slack rich_text conversion
          const htmlContent = (msg.message_html && !isOverBoldedHtml(msg.message_html, msg.message)) ? msg.message_html : '';
          const formattedContent = formatMessageContent(msg.message || '');

          // If we have rich HTML from the parser, check if it's simple enough to render inline
          if (htmlContent) {
            // Simple HTML: just text in basic wrapper divs with no complex formatting
            const isSimple = (() => {
              const t = document.createElement('div');
              t.innerHTML = htmlContent;
              const text = t.textContent || '';
              // Simple if: short text, no tables, no images, no iframes, few style attributes, no signatures
              if (text.length > 500) return false;
              if (t.querySelectorAll('table, img, iframe, object, embed').length > 0) return false;
              if (t.querySelectorAll('[class*="Signature"], [class*="signature"], .elementToProof').length > 1) return false;
              if ((htmlContent.match(/style="/g) || []).length > 5) return false;
              if (t.querySelectorAll('hr').length > 0) return false;
              if (t.querySelectorAll('blockquote').length > 0) return false;
              return true;
            })();
            if (isSimple) {
              msgContent.innerHTML = typeof sanitizeHtml === 'function' ? sanitizeHtml(htmlContent) : htmlContent;
            } else {
              msgContent.dataset.emailIframe = 'true';
              msgContent.dataset.rawHtml = htmlContent;
            }
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
              const tooltipText = users.length > 0 ? users.map(u => typeof u === 'object' ? (u.name || u.user_id) : u).join(', ') : `${count} reaction${count > 1 ? 's' : ''}`;

              const pill = document.createElement('span');
              pill.className = 'reaction-pill';
              pill.style.cssText = `display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: ${isDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border: 1px solid ${isDarkMode ? 'rgba(102, 126, 234, 0.25)' : 'rgba(102, 126, 234, 0.15)'}; color: ${isDarkMode ? '#b0b0b0' : '#555'}; cursor: default; position: relative;`;
              // For unresolved emoji names, style them smaller so they fit in the pill
              const isResolved = !emojiUnicode.startsWith(':');
              const emojiDisplay = isResolved
                ? `<span style="font-size: 14px;">${emojiUnicode}</span>`
                : `<span style="font-size: 11px; opacity: 0.8;">${emojiUnicode}</span>`;
              pill.innerHTML = `${emojiDisplay}${count > 1 ? `<span style="font-size: 11px; font-weight: 500;">${count}</span>` : ''}`;

              // Custom styled tooltip — wraps for long participant lists
              const tooltip = document.createElement('div');
              tooltip.style.cssText = `display: none; position: absolute; bottom: calc(100% + 6px); left: 50%; transform: translateX(-50%); padding: 6px 10px; border-radius: 8px; font-size: 11px; line-height: 1.4; max-width: 280px; white-space: normal; z-index: 10001; pointer-events: none; background: ${isDarkMode ? '#2a2a2a' : '#333'}; color: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.25);`;
              tooltip.innerHTML = formatReactionTooltip(tooltipText);
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
            threadDiv.style.cssText = `margin-left: 40px; margin-top: 4px; display: flex; align-items: center; gap: 6px; cursor: pointer; flex-wrap: wrap; ${isCollapsed ? 'display: none;' : ''}`;

            const replyUsers = msg.thread.reply_users || [];
            const replyCount = msg.thread.reply_count;
            const latestReply = msg.thread.latest_reply || '';

            // Mini avatars for reply users (max 3)
            let avatarsHtml = '';
            const showUsers = replyUsers.slice(0, 3);
            showUsers.forEach((user, i) => {
              const userName = typeof user === 'object' ? (user.name || 'U') : (typeof user === 'string' ? user : 'U');
              const initial = userName.charAt(0).toUpperCase();
              const userChipColor = getUserChipColor(typeof user === 'object' ? (user.user_id || userName) : userName);
              avatarsHtml += `<div style="width: 20px; height: 20px; border-radius: 50%; background: ${userChipColor}; display: flex; align-items: center; justify-content: center; color: white; font-size: 9px; font-weight: 600; margin-left: ${i > 0 ? '-4px' : '0'}; border: 1.5px solid ${isDarkMode ? '#1f2940' : 'white'}; position: relative; z-index: ${3 - i};">${initial}</div>`;
            });

            const latestReplyText = latestReply ? ` · Last reply ${formatTimeAgoFresh(latestReply)}` : '';

            threadDiv.innerHTML = `
              <div style="display: flex; align-items: center;">${avatarsHtml}</div>
              <span class="thread-replies-link" style="font-size: 12px; color: #667eea; font-weight: 600; cursor: pointer; text-decoration: underline dotted;">${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}</span>
              <span style="font-size: 11px; color: ${isDarkMode ? '#666' : '#95a5a6'};">${latestReplyText}</span>
            `;
            // Extract channel from todo's message_link for thread panel
            const msgChannelId = extractSlackChannelId(todo.message_link || '');
            threadDiv.addEventListener('click', () => openThreadPanel(slider, msgChannelId, todo.participant_text || '', msg.ts || msg.time, msg.message_from || 'Unknown', msg.message || '', isDarkMode));

            msgDiv.appendChild(threadDiv);
          }

          // "Reply" label - shown on hover; hidden for slack_thread type (use "X replies" link), messages with replies, and Gmail threads
          const _isSlackThread = (todo.type === 'slack_thread');
          const _msgHasReplies = msg.thread && msg.thread.reply_count > 0;
          const _isGmailTask = (todo.message_link || '').includes('mail.google.com');
          if (!_isSlackThread && !_msgHasReplies && !_isGmailTask) {
            const replyIconBtnTs = document.createElement('div');
            replyIconBtnTs.className = 'msg-reply-btn';
            replyIconBtnTs.title = 'Reply in thread';
            replyIconBtnTs.textContent = 'Reply';
            replyIconBtnTs.style.cssText = `position: absolute; bottom: 6px; right: 8px; font-size: 11px; font-weight: 600; color: #667eea; opacity: 0; cursor: pointer; transition: opacity 0.1s; z-index: 2; background: ${isDarkMode ? '#1f2940' : 'white'}; border: 1px solid ${isDarkMode ? 'rgba(102,126,234,0.3)' : 'rgba(102,126,234,0.25)'}; border-radius: 6px; padding: 2px 8px; line-height: 1.6;`;
            const _msgChannelId = extractSlackChannelId(todo.message_link || '');
            const _msgChannelName = todo.participant_text || '';
            replyIconBtnTs.addEventListener('click', (e) => {
              e.stopPropagation();
              document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
              document.querySelectorAll('.transcript-slider-backdrop').forEach(b => b.remove());
              isTranscriptSliderOpen = false;
              openReplySlider(_msgChannelId, _msgChannelName, msg, isDarkMode);
            });
            msgDiv.style.position = 'relative';
            msgDiv.appendChild(replyIconBtnTs);
            msgDiv.addEventListener('mouseenter', () => replyIconBtnTs.style.opacity = '1');
            msgDiv.addEventListener('mouseleave', () => replyIconBtnTs.style.opacity = '0');
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

        // Scroll to show the source message highlight, or the latest (last) message
        setTimeout(() => {
          const highlightedMessage = messagesContainer.querySelector('.source-message-highlight');
          if (highlightedMessage && sourceMessageIndex > 0) {
            highlightedMessage.scrollIntoView({ behavior: 'auto', block: 'center' });
          } else {
            const lastMessage = messagesContainer.lastElementChild;
            if (lastMessage) {
              lastMessage.scrollIntoView({ behavior: 'auto', block: 'start' });
            }
          }
          // Render complex email HTML in iframes now that elements are in DOM
          messagesContainer.querySelectorAll('[data-email-iframe="true"]').forEach(el => {
            renderEmailInIframe(el.dataset.rawHtml, el, isDarkMode);
            delete el.dataset.emailIframe;
            delete el.dataset.rawHtml;
          });
          // CSP-safe avatar image fallback (inline onerror is blocked in MV3)
          messagesContainer.querySelectorAll('img.transcript-avatar-img').forEach(img => {
            img.addEventListener('error', function() {
              this.style.display = 'none';
              const fallback = this.nextElementSibling;
              if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
                fallback.style.display = 'flex';
              }
            });
            // Handle already-failed images (cached failures)
            if (img.complete && img.naturalWidth === 0) {
              img.style.display = 'none';
              const fallback = img.nextElementSibling;
              if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
                fallback.style.display = 'flex';
              }
            }
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

    // Rich text formatting for contenteditable reply input
    setupRichTextFormatting(replyInput);

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

        // Prepare thread data — only from source message onwards, capped at 5
        let filteredMessages = messages;
        const srcIdx = slider._sourceMessageIndex ?? -1;
        if (srcIdx >= 0) {
          filteredMessages = messages.slice(srcIdx);
        }
        if (filteredMessages.length > 5) {
          filteredMessages = filteredMessages.slice(-5);
        }
        const threadConversation = filteredMessages.map(msg => ({
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

        oracleConversationHistory.push({ role: 'user', content: userMessage || '' });

        const payload = {
          session_id: sessionId,
          timestamp: new Date().toISOString(),
          source: isFollowUp ? 'oracle-transcript-followup' : 'oracle-transcript',
          user_id: userId,
          context: {
            task_title: currentTodo.task_title || '',
            task_name: currentTodo.task_name || '',
            message_link: currentTodo.message_link || '',
            ...(isFollowUp ? {} : {
              conversation: threadConversation.map(msg => ({
                person: msg.role,
                message: msg.content,
                timestamp: msg.timestamp
              }))
            })
          }
        };

        // For follow-ups, send user message and conversation history at top level
        if (isFollowUp) {
          payload.message = userMessage;
          payload.conversation = oracleConversationHistory;
        }

        // Function to convert URLs to hyperlinks with icons
        const formatOracleResponse = (text) => {
          // Extract markdown tables BEFORE escapeHtml
          const tablePlaceholders = [];
          let preProcessed = text.replace(/(?:^|\n)((?:\|[^\n]+\|\s*\n){2,})/gm, (match, tableBlock) => {
            const lines = tableBlock.trim().split('\n').filter(l => l.trim());
            if (lines.length < 2) return match;
            const isSeparator = (line) => /^\|[\s\-:]+(\|[\s\-:]+)+\|?\s*$/.test(line.trim());
            let headerLine, dataLines;
            if (isSeparator(lines[1])) {
              headerLine = lines[0];
              dataLines = lines.slice(2);
            } else {
              headerLine = null;
              dataLines = lines;
            }
            const parseRow = (line) => line.split('|').slice(1, -1).map(c => c.trim());
            const borderColor = isDarkMode ? 'rgba(255,255,255,0.1)' : 'rgba(102,126,234,0.15)';
            const headerBg = isDarkMode ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.08)';
            const stripeBg = isDarkMode ? 'rgba(255,255,255,0.03)' : 'rgba(102,126,234,0.03)';
            let html = '<div style="overflow-x:auto;margin:8px 0;"><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid ' + borderColor + ';border-radius:8px;overflow:hidden;">';
            if (headerLine) {
              const cells = parseRow(headerLine);
              html += '<thead><tr style="background:' + headerBg + ';">' + cells.map(c => '<th style="padding:8px 12px;text-align:left;font-weight:600;border-bottom:2px solid ' + borderColor + ';white-space:nowrap;">' + escapeHtml(c) + '</th>').join('') + '</tr></thead>';
            }
            html += '<tbody>';
            dataLines.forEach((line, i) => {
              const cells = parseRow(line);
              const bg = i % 2 === 1 ? stripeBg : 'transparent';
              html += '<tr style="background:' + bg + ';">' + cells.map(c => '<td style="padding:6px 12px;border-bottom:1px solid ' + borderColor + ';">' + escapeHtml(c) + '</td>').join('') + '</tr>';
            });
            html += '</tbody></table></div>';
            const placeholder = '%%ORACLE_TABLE_' + tablePlaceholders.length + '%%';
            tablePlaceholders.push(html);
            return '\n' + placeholder + '\n';
          });

          let formatted = escapeHtml(preProcessed);

          // Restore table placeholders
          tablePlaceholders.forEach((html, i) => {
            formatted = formatted.replace('%%ORACLE_TABLE_' + i + '%%', html);
          });

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
          // Native n8n webhook streaming (V38) — parses NDJSON stream
          let fullResponseText = '', streamStarted = false;
          const controller = new AbortController();
          const streamTimeout = setTimeout(() => {
            controller.abort();
            if (!streamStarted) {
              const loadEl = oracleContent.querySelector('.oracle-loading');
              if (loadEl) loadEl.innerHTML = '<div style="padding:16px;color:#e74c3c;text-align:center;"><span style="font-size:24px;">⏱</span><div>Response timed out</div></div>';
            }
          }, 120000);

          const response = await fetch('https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
            signal: controller.signal
          });

          if (!response.ok) throw new Error('HTTP ' + response.status);

          const reader = response.body.getReader();
          const decoder = new TextDecoder();
          let buffer = '';

          try {
            while (true) {
              const { done, value } = await reader.read();
              if (done) break;

              buffer += decoder.decode(value, { stream: true });
              const lines = buffer.split('\n');
              buffer = lines.pop() || '';

              for (const line of lines) {
                if (!line.trim()) continue;
                try {
                  const parsed = JSON.parse(line);
                  if (parsed.type === 'item' && parsed.content && parsed.metadata?.nodeName !== 'Respond to Webhook') {
                    if (!streamStarted) {
                      streamStarted = true;
                      const loadEl = oracleContent.querySelector('.oracle-loading');
                      if (loadEl) loadEl.remove();
                      const responseDiv = document.createElement('div');
                      responseDiv.className = 'oracle-response';
                      responseDiv.style.cssText = 'white-space:pre-wrap;word-wrap:break-word;margin-right:auto;max-width:95%;padding:10px 14px;background:' + (isDarkMode ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.04)') + ';border-radius:10px;border-top-left-radius:4px;margin-bottom:8px;';
                      oracleContent.appendChild(responseDiv);
                    }
                    fullResponseText += parsed.content;
                    const responseDiv = oracleContent.querySelector('.oracle-response:last-of-type');
                    if (responseDiv) responseDiv.innerHTML = formatOracleResponse(fullResponseText);
                    // Auto-scroll only if the bottom of the response is still within the visible area
                    const atBottom = oracleContent.scrollHeight - oracleContent.scrollTop <= oracleContent.clientHeight + 80;
                    if (atBottom) oracleContent.scrollTop = oracleContent.scrollHeight;
                  }
                } catch { /* skip unparseable lines */ }
              }
            }
            // Process remaining buffer
            if (buffer.trim()) {
              try {
                const parsed = JSON.parse(buffer);
                if (parsed.type === 'item' && parsed.content && parsed.metadata?.nodeName !== 'Respond to Webhook') {
                  fullResponseText += parsed.content;
                }
              } catch { /* skip */ }
            }
          } finally {
            clearTimeout(streamTimeout);
            reader.releaseLock();
          }

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
          const finalDiv = oracleContent.querySelector('.oracle-response:last-of-type');
          if (finalDiv) finalDiv.innerHTML = formatOracleResponse(fullResponseText);
          addOracleCopyButtons(fullResponseText);
          oracleContent.scrollTop = oracleContent.scrollHeight;
        } catch (error) {
          console.error('Oracle request failed:', error);
          const loadEl = oracleContent.querySelector('.oracle-loading');
          if (loadEl) loadEl.innerHTML = '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;color:#e74c3c;"><span style="font-size:24px;">⚠️</span><div>Failed to get response from Oracle</div><div style="font-size:12px;color:' + (isDarkMode ? '#888' : '#95a5a6') + ';">' + escapeHtml(error.message) + '</div></div>';
        }
      };

      // Oracle button click - initial thread analysis
      oracleBtn.addEventListener('click', async () => {
        // Open assistant slider with "Reading the conversation..." state
        if (!isChatSliderOpen) {
          window.OracleAssistant.showChatSlider({
            mode: 'fullscreen',
            showReadingState: true,
            onClose: () => { isChatSliderOpen = false; }
          });
          isChatSliderOpen = true;
        } else {
          // Slider already open — inject reading state onto existing overlay
          const existingOverlay = document.querySelector('.chat-slider-overlay');
          if (existingOverlay) {
            const isDark = document.body.classList.contains('dark-mode');
            const msgContainer = existingOverlay.querySelector('.chat-messages');
            if (msgContainer) {
              // Remove any stale reading state
              msgContainer.querySelectorAll('.oracle-reading-state').forEach(el => el.remove());
              const readingDiv = document.createElement('div');
              readingDiv.className = 'oracle-reading-state';
              readingDiv.style.cssText = 'display:flex;align-items:center;gap:10px;padding:16px;';
              readingDiv.innerHTML = '<div style="width:28px;height:28px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;flex-shrink:0;"></div><div style="color:' + (isDark ? '#888' : '#7f8c8d') + ';font-size:13px;">Reading the conversation...</div>';
              msgContainer.appendChild(readingDiv);
              msgContainer.scrollTop = msgContainer.scrollHeight;
            }
            // Register resolver on existing overlay so Ably message can clear it
            existingOverlay._oracleResolveQuery = (queryText) => {
              const readEl = msgContainer?.querySelector('.oracle-reading-state');
              if (readEl) readEl.remove();
              const chatInput = existingOverlay.querySelector('.chat-input');
              const sendBtn = existingOverlay.querySelector('.chat-send-btn');
              if (chatInput) { chatInput.setAttribute('contenteditable', 'true'); chatInput.style.opacity = '1'; }
              if (sendBtn) { sendBtn.disabled = false; sendBtn.style.opacity = '1'; }
              // Show user bubble and auto-send
              if (msgContainer) {
                const userMsgDiv = document.createElement('div');
                userMsgDiv.style.cssText = 'display:flex;justify-content:flex-end;';
                const userBubble = document.createElement('div');
                userBubble.style.cssText = 'max-width:80%;padding:12px 16px;background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:16px 16px 4px 16px;font-size:14px;line-height:1.5;';
                userBubble.textContent = queryText;
                userMsgDiv.appendChild(userBubble);
                msgContainer.appendChild(userMsgDiv);
              }
              _skipNextUserBubble = true;
              if (chatInput) chatInput.textContent = queryText;
              setTimeout(() => {
                const sendFn = existingOverlay.querySelector('.chat-send-btn');
                if (sendFn) sendFn.click();
              }, 100);
            };
          }
        }

        // Fire webhook to n8n for query extraction only (no AI Agent)
        const allMessages = slider.transcriptMessages || [];
        const currentTodo = slider.transcriptTodo || {};
        
        // Only send messages from source message onwards, capped at 5
        let contextMessages;
        const srcMsgIdx = slider._sourceMessageIndex ?? -1;
        if (srcMsgIdx >= 0) {
          contextMessages = allMessages.slice(srcMsgIdx);
        } else {
          contextMessages = allMessages;
        }
        // Cap at last 5 messages
        if (contextMessages.length > 5) {
          contextMessages = contextMessages.slice(-5);
        }
        let userId = window.Oracle?.state?.userData?.userId;
        if (!userId) {
          try { const r = await chrome.storage.local.get(['oracle_user_data']); userId = r?.oracle_user_data?.userId || null; } catch { userId = null; }
        }
        const payload = {
          session_id: 'transcript_' + todoId + '_' + Date.now(),
          timestamp: new Date().toISOString(),
          source: 'oracle-transcript',
          user_id: userId,
          context: {
            task_title: currentTodo.task_title || '',
            task_name: currentTodo.task_name || '',
            message_link: currentTodo.message_link || '',
            conversation: contextMessages.map(msg => ({
              person: msg.user || msg.sender || msg.message_from || 'unknown',
              message: msg.text || msg.body || msg.message || '',
              timestamp: msg.time || msg.timestamp || msg.ts || ''
            }))
          }
        };
        try {
          await fetch('https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
          });
        } catch (err) { console.error('Transcript webhook error:', err); }
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

    // Keyboard shortcut: Cmd/Ctrl+Enter → send (other formatting handled by setupRichTextFormatting)
    replyInput.addEventListener('keydown', (e) => {
      // Cmd/Ctrl+Enter → send
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        const sendBtn = slider.querySelector('.transcript-send-btn');
        if (sendBtn && !sendBtn.disabled) sendBtn.click();
        return;
      }
      // Plain Enter → if inside a list, let setupRichTextFormatting handle it (list continuation/exit);
      // otherwise insert line break (not <div>)
      // Skip if already handled by setupRichTextFormatting (e.g. empty list item exit)
      if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.shiftKey && !e.defaultPrevented) {
        const sel = window.getSelection();
        const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
        const inList = node && (node.nodeType === Node.TEXT_NODE
          ? node.parentElement?.closest('ol, ul')
          : node.closest?.('ol, ul'));
        if (!inList) {
          e.preventDefault();
          document.execCommand('insertLineBreak');
        }
        // If inside a list, don't preventDefault — let browser default create next <li>
        return;
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

    // @Mention functionality — now handled by shared setupMentionHandler
    setupMentionHandler(replyInput, replySection);
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

    // Function to convert reply text to Slack format (replace @Name with <@SLACK_ID>, convert formatting)
    const getSlackFormattedText = () => {
      return convertContentEditableToSlackMrkdwn(replyInput);
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

          // Normalize field names from API and filter to people only (exclude channels for Gmail CC/BCC)
          data = data.filter(r => r.type === 'employee' || r.type === 'Direct Message' || (!r.type && r.user_email_ID));
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

    // Note: mention insertion auto-CC is handled within setupMentionHandler.
    // The CC logic for email tasks hooks into getMentionedUsersFromInput() at send time.

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

      // Optimistic: show message immediately
      const replyHtmlForOptimistic = replyHtml;
      appendOptimisticMessage(messagesContainer, replyText, replyHtmlForOptimistic, isDarkMode);
      replyInput.innerHTML = '';

      sendBtn.disabled = true;
      sendBtn.innerHTML = '<span style="font-size: 14px;">⏳</span>';

      // Get mentioned users from the input
      const mentionedUsersInReply = getMentionedUsersFromInput();
      console.log('Mentioned users:', mentionedUsersInReply);
      console.log('Slack formatted text:', replyTextSlack);

      try {
        // Determine the reply action based on task type
        const isSlackThread = todo.type === 'slack_thread';
        const isSlackMessage = todo.type && todo.type !== 'slack_thread' && !isDriveTask && !todo.type.includes('gmail') && !todo.type.includes('drive') && (todo.message_link || '').includes('slack.com');
        const replyAction = isDriveTask ? 'reply_drive_comment' : (isSlackMessage ? 'reply_to_channel' : 'reply_to_thread');
        const slackChannelIdForReply = isSlackMessage ? extractSlackChannelId(todo.message_link || '') : undefined;

        const payload = createAuthenticatedPayload({
          action: replyAction,
          todo_id: todoId,
          message_link: todo.message_link,
          channel_id: slackChannelIdForReply || undefined,
          reply_text: isDriveTask ? replyText : replyTextSlack,
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
          // Make optimistic message fully opaque on success
          const optMsg = messagesContainer.querySelector('.optimistic-message:last-child');
          if (optMsg) optMsg.style.opacity = '1';

          // Clear mentioned users (tracked via DOM mention-tags, no separate array needed)

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
                  const responseData = Array.isArray(data) ? data[0] : data;
                  const transcript = responseData?.transcript || [];
                  const refreshUserDirectory = responseData?.user_directory || {};

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

                  // Re-render messages (full rendering with reactions, threads, avatars)
                  if (messages.length > 0) {
                    messagesContainer.innerHTML = '';
                    // Check current dark mode state for refresh
                    const refreshDarkMode = document.body.classList.contains('dark-mode');
                    messages.forEach(msg => {
                      const msgDiv = document.createElement('div');
                      msgDiv.style.cssText = 'display: flex; flex-direction: column; gap: 6px;';

                      const msgHeader = document.createElement('div');
                      msgHeader.style.cssText = 'display: flex; align-items: center; gap: 8px;';

                      // Generate consistent color from user_id for name chips
                      const chipColor = getUserChipColor(msg.user_id || msg.message_from || 'Unknown');
                      const avatarUrl = msg.avatar_url || (refreshUserDirectory[msg.user_id]?.avatar_url);
                      const nameParts = (msg.message_from || 'U').trim().split(/\s+/);
                      const initials = nameParts.length > 1
                        ? (nameParts[0].charAt(0) + nameParts[nameParts.length - 1].charAt(0)).toUpperCase()
                        : nameParts[0].charAt(0).toUpperCase();
                      const initialsFallback = `<div style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`;
                      const avatarHtml = avatarUrl
                        ? `<img src="${escapeHtml(avatarUrl)}" class="transcript-avatar-img" data-fallback-initials="${initials}" data-fallback-color="${chipColor}" data-fallback-size="${initials.length > 1 ? '11' : '13'}" style="width: 32px; height: 32px; border-radius: 50%; object-fit: cover; flex-shrink: 0;" /><div class="transcript-avatar-fallback" style="width: 32px; height: 32px; background: ${chipColor}; border-radius: 50%; display: none; align-items: center; justify-content: center; color: white; font-size: ${initials.length > 1 ? '11' : '13'}px; font-weight: 600; flex-shrink: 0;">${initials}</div>`
                        : initialsFallback;

                      msgHeader.innerHTML = `
                        ${avatarHtml}
                        <div style="flex: 1; min-width: 0;">
                          <div style="font-weight: 600; font-size: 13px; color: ${refreshDarkMode ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(msg.message_from || 'Unknown')}</div>
                          <div style="font-size: 11px; color: ${refreshDarkMode ? '#666' : '#95a5a6'};">${escapeHtml(formatTimeAgoFresh(msg.time || ''))}</div>
                        </div>
                      `;

                      const msgContent = document.createElement('div');
                      msgContent.style.cssText = `margin-left: 40px; padding: 12px 16px; background: ${refreshDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border-radius: 12px; border-top-left-radius: 4px; font-weight: 400; font-size: 14px; color: ${refreshDarkMode ? '#e8e8e8' : '#2c3e50'}; line-height: 1.6; word-wrap: break-word; overflow-wrap: break-word; word-break: break-word; overflow-x: hidden; overflow-y: visible; max-width: 100%;`;
                      // Skip message_html if it has over-bolded content from Slack rich_text conversion
                      const htmlContent = (msg.message_html && !isOverBoldedHtml(msg.message_html, msg.message)) ? msg.message_html : '';
                      if (htmlContent) {
                        const isSimple = (() => {
                          const t = document.createElement('div');
                          t.innerHTML = htmlContent;
                          const text = t.textContent || '';
                          if (text.length > 500) return false;
                          if (t.querySelectorAll('table, img, iframe, object, embed').length > 0) return false;
                          if (t.querySelectorAll('[class*="Signature"], [class*="signature"], .elementToProof').length > 1) return false;
                          if ((htmlContent.match(/style="/g) || []).length > 5) return false;
                          if (t.querySelectorAll('hr, blockquote').length > 0) return false;
                          return true;
                        })();
                        if (isSimple) {
                          msgContent.innerHTML = typeof sanitizeHtml === 'function' ? sanitizeHtml(htmlContent) : htmlContent;
                        } else {
                          msgContent.dataset.emailIframe = 'true';
                          msgContent.dataset.rawHtml = htmlContent;
                        }
                      } else if (isComplexEmailHtml(msg.message || '')) {
                        msgContent.dataset.emailIframe = 'true';
                        msgContent.dataset.rawHtml = msg.message;
                      } else {
                        msgContent.innerHTML = formatMessageContent(msg.message || '');
                      }

                      msgDiv.appendChild(msgHeader);
                      msgDiv.appendChild(msgContent);

                      // Render attachments if present
                      if (msg.attachments && msg.attachments.length > 0) {
                        const validAttachments = msg.attachments.filter(att => {
                          if (att.type === 'link' && (!att.name || att.name === 'Attachment')) return false;
                          if (!att.name && !att.url && !att.text) return false;
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

                      // Render reactions if present
                      if (msg.reactions && msg.reactions.length > 0) {
                        const reactionsDiv = document.createElement('div');
                        reactionsDiv.className = 'transcript-message-reactions';
                        reactionsDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; flex-wrap: wrap; gap: 4px;';

                        msg.reactions.forEach(reaction => {
                          const emojiName = reaction.emoji || reaction.name || '';
                          const count = reaction.count || 1;
                          const users = reaction.users || [];

                          const emojiUnicode = convertSlackEmoji(emojiName);
                          const tooltipText = users.length > 0 ? users.map(u => typeof u === 'object' ? (u.name || u.user_id) : u).join(', ') : `${count} reaction${count > 1 ? 's' : ''}`;

                          const pill = document.createElement('span');
                          pill.className = 'reaction-pill';
                          pill.style.cssText = `display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 12px; font-size: 12px; background: ${refreshDarkMode ? 'rgba(102, 126, 234, 0.15)' : 'rgba(102, 126, 234, 0.08)'}; border: 1px solid ${refreshDarkMode ? 'rgba(102, 126, 234, 0.25)' : 'rgba(102, 126, 234, 0.15)'}; color: ${refreshDarkMode ? '#b0b0b0' : '#555'}; cursor: default; position: relative;`;
                          pill.innerHTML = `<span style="font-size: 14px;">${emojiUnicode}</span>${count > 1 ? `<span style="font-size: 11px; font-weight: 500;">${count}</span>` : ''}`;

                          const tooltip = document.createElement('div');
                          tooltip.style.cssText = `display: none; position: absolute; bottom: calc(100% + 6px); left: 50%; transform: translateX(-50%); padding: 6px 10px; border-radius: 8px; font-size: 11px; line-height: 1.4; max-width: 280px; white-space: normal; z-index: 10001; pointer-events: none; background: ${refreshDarkMode ? '#2a2a2a' : '#333'}; color: #fff; box-shadow: 0 2px 8px rgba(0,0,0,0.25);`;
                          tooltip.innerHTML = formatReactionTooltip(tooltipText);
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
                        threadDiv.style.cssText = 'margin-left: 40px; margin-top: 4px; display: flex; align-items: center; gap: 6px;';

                        const replyUsers = msg.thread.reply_users || [];
                        const replyCount = msg.thread.reply_count;
                        const latestReply = msg.thread.latest_reply || '';

                        let avatarsHtml = '';
                        const showUsers = replyUsers.slice(0, 3);
                        showUsers.forEach((user, i) => {
                          const userName = typeof user === 'object' ? (user.name || 'U') : (typeof user === 'string' ? user : 'U');
                          const initial = userName.charAt(0).toUpperCase();
                          const userChipColor = getUserChipColor(typeof user === 'object' ? (user.user_id || userName) : userName);
                          avatarsHtml += `<div style="width: 20px; height: 20px; border-radius: 50%; background: ${userChipColor}; display: flex; align-items: center; justify-content: center; color: white; font-size: 9px; font-weight: 600; margin-left: ${i > 0 ? '-4px' : '0'}; border: 1.5px solid ${refreshDarkMode ? '#1f2940' : 'white'}; position: relative; z-index: ${3 - i};">${initial}</div>`;
                        });

                        const latestReplyText = latestReply ? ` · Last reply ${formatTimeAgoFresh(latestReply)}` : '';
                        threadDiv.innerHTML = `
                          <div style="display: flex; align-items: center;">${avatarsHtml}</div>
                          <span style="font-size: 12px; color: #667eea; font-weight: 600; cursor: default;">${replyCount} ${replyCount === 1 ? 'reply' : 'replies'}</span>
                          <span style="font-size: 11px; color: ${refreshDarkMode ? '#666' : '#95a5a6'};">${latestReplyText}</span>
                        `;

                        msgDiv.appendChild(threadDiv);
                      }

                      messagesContainer.appendChild(msgDiv);
                    });

                    // Post-render: iframes, avatar fallbacks, scroll
                    setTimeout(() => {
                      messagesContainer.querySelectorAll('[data-email-iframe="true"]').forEach(el => {
                        renderEmailInIframe(el.dataset.rawHtml, el, refreshDarkMode);
                        delete el.dataset.emailIframe;
                        delete el.dataset.rawHtml;
                      });
                      messagesContainer.querySelectorAll('img.transcript-avatar-img').forEach(img => {
                        img.addEventListener('error', function() {
                          this.style.display = 'none';
                          const fallback = this.nextElementSibling;
                          if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
                            fallback.style.display = 'flex';
                          }
                        });
                        if (img.complete && img.naturalWidth === 0) {
                          img.style.display = 'none';
                          const fallback = img.nextElementSibling;
                          if (fallback && fallback.classList.contains('transcript-avatar-fallback')) {
                            fallback.style.display = 'flex';
                          }
                        }
                      });
                    }, 50);

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
        // Remove optimistic message on failure and restore input
        const optMsg = messagesContainer.querySelector('.optimistic-message:last-child');
        if (optMsg) { optMsg.style.opacity = '0'; setTimeout(() => optMsg.remove(), 300); }
        replyInput.innerHTML = replyHtmlForOptimistic;
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

    // Strip trailing numbers: e.g. "welcome2" → "welcome", "tada2" → "tada"
    const numStripped = name.replace(/\d+$/, '');
    if (numStripped && numStripped !== name && emojiMap[numStripped]) return emojiMap[numStripped];

    // Common custom Slack emoji fallbacks
    const customFallbacks = {
      'welcome':'👋','thanks':'🙏','thankyou':'🙏','thank_you':'🙏',
      'congrats':'🎉','celebrate':'🎉','yay':'🎉','lgtm':'👍',
      'shipit':'🚀','ship':'🚀','approved':'✅','done':'✅',
      'ack':'👍','noted':'📝','alert':'🚨','hi':'👋','hello':'👋',
      'great':'🔥','awesome':'🔥','nice':'👍','good':'👍','cool':'😎',
      'blob_wave':'👋','blob_clap':'👏','blob_thumbsup':'👍',
      'blob_heart':'❤️','blob_thinking':'🤔','blob_eyes':'👀',
      'parrot':'🦜','meow':'🐱','doge':'🐕',
      // Animated / workspace custom emoji fallbacks
      'gifire':'🔥','fire-but-animated':'🔥','fire_gif':'🔥','fireflame':'🔥',
      'parrotdoge':'🦜','partydoge':'🎉','partyparrot':'🦜','party_parrot':'🦜',
      'fastparrot':'🦜','slowparrot':'🦜','shuffleparrot':'🦜','congaparrot':'🦜',
      'blobdance':'💃','blob_dance':'💃','blobjam':'🎵','catjam':'🐱',
      'meow_party':'🐱','nyan':'🐱','nyancat':'🐱',
      'dancingbanana':'🍌','bananadance':'🍌',
      'loadingdots':'⏳','loading':'⏳','spinner':'⏳',
      'stonks':'📈','notstonks':'📉',
      'this':'👆','that':'👆','upvote':'👍','downvote':'👎',
      'applause':'👏','clapping':'👏','slow_clap':'👏','clap':'👏',
      'bow':'🙇','salute':'🫡','facepalm':'🤦','shrug':'🤷',
    };
    if (customFallbacks[name]) return customFallbacks[name];
    if (numStripped && customFallbacks[numStripped]) return customFallbacks[numStripped];

    return `:${name}:`;
  }

  // formatReactionTooltip — wraps at commas between names, not within names
  function formatReactionTooltip(text) {
    if (!text) return '';
    const names = text.split(',').map(n => n.trim());
    return names.map((name, i) => 
      `<span style="white-space:nowrap;display:inline-block">${escapeHtml(name)}${i < names.length - 1 ? ',' : ''}</span>`
    ).join(' ');
  }

  // getUserChipColor — generate a consistent color from a user identifier
  // Uses a hash of the string to pick from a curated palette
  function getUserChipColor(identifier) {
    const palette = [
      'linear-gradient(135deg, #667eea, #764ba2)',  // purple-blue
      'linear-gradient(135deg, #f093fb, #f5576c)',  // pink
      'linear-gradient(135deg, #4facfe, #00f2fe)',  // cyan
      'linear-gradient(135deg, #43e97b, #38f9d7)',  // green
      'linear-gradient(135deg, #fa709a, #fee140)',  // coral-yellow
      'linear-gradient(135deg, #a18cd1, #fbc2eb)',  // lavender
      'linear-gradient(135deg, #fccb90, #d57eeb)',  // peach-purple
      'linear-gradient(135deg, #e0c3fc, #8ec5fc)',  // soft purple-blue
      'linear-gradient(135deg, #f6d365, #fda085)',  // warm sunset
      'linear-gradient(135deg, #89f7fe, #66a6ff)',  // sky blue
      'linear-gradient(135deg, #fddb92, #d1fdff)',  // cream-mint
      'linear-gradient(135deg, #c471f5, #fa71cd)',  // magenta
    ];
    let hash = 0;
    const str = String(identifier || 'U');
    for (let i = 0; i < str.length; i++) {
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash = hash & hash; // Convert to 32-bit integer
    }
    return palette[Math.abs(hash) % palette.length];
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

  // Delay initial data fetch by 3 seconds to avoid unnecessary n8n webhook calls
  // when users quickly open and switch away from new tabs.
  // Show a splash loader with rotating motivational quotes during the wait.
  const INITIAL_LOAD_DELAY = 3000;
  let initialLoadTimer = null;

  // --- Splash screen ---
  const splashQuotes = [
    { text: "Brewing your priorities...", icon: "☕" },
    { text: "Syncing the universe of tasks...", icon: "🌌" },
    { text: "Aligning your day for greatness...", icon: "✨" },
    { text: "Gathering intel from all channels...", icon: "📡" },
    { text: "Almost there, deep breaths...", icon: "🧘" },
    { text: "Good things take a moment...", icon: "⏳" },
    { text: "Loading your second brain...", icon: "🧠" },
    { text: "Connecting the dots for you...", icon: "🔗" },
    { text: "You're going to crush it today...", icon: "💪" },
    { text: "Warming up the engines...", icon: "🚀" },
    { text: "Scanning your horizon...", icon: "🔭" },
    { text: "Tuning into your workflow...", icon: "📻" },
    { text: "Assembling the big picture...", icon: "🧩" },
    { text: "Calibrating your command center...", icon: "🎛️" },
    { text: "Polishing your crystal ball...", icon: "🔮" },
    { text: "Summoning your daily briefing...", icon: "📋" },
    { text: "Crunching the latest signals...", icon: "📊" },
    { text: "Fetching your superpowers...", icon: "🦸" },
    { text: "Spinning up your co-pilot...", icon: "🛩️" },
    { text: "Charting the course ahead...", icon: "🗺️" },
    { text: "Powering up the Oracle...", icon: "⚡" },
    { text: "Queuing up what matters most...", icon: "🎯" },
    { text: "Locking in your focus zone...", icon: "🔒" },
    { text: "Mapping today's priorities...", icon: "📍" },
    { text: "Decoding your schedule...", icon: "🗓️" },
    { text: "Sorting signal from noise...", icon: "📶" },
    { text: "Revving up your productivity...", icon: "🏎️" },
    { text: "Sharpening your edge...", icon: "🪓" },
    { text: "Aligning the stars for you...", icon: "🌟" },
    { text: "Lighting up your dashboard...", icon: "💡" },
    { text: "Unrolling today's scroll...", icon: "📜" },
    { text: "Warming up your launchpad...", icon: "🛫" },
    { text: "Downloading clarity...", icon: "💎" },
    { text: "Preparing your mission briefing...", icon: "🎖️" },
    { text: "Building your battle plan...", icon: "⚔️" },
    { text: "Tuning the instruments...", icon: "🎵" },
    { text: "Reticulating your splines...", icon: "🌀" },
    { text: "Consulting the oracle...", icon: "🏛️" },
    { text: "Harnessing the data streams...", icon: "🌊" },
    { text: "Compiling your advantage...", icon: "🏆" },
    { text: "Orchestrating your workflow...", icon: "🎼" },
    { text: "Getting your ducks in a row...", icon: "🦆" },
    { text: "Fueling up for the day...", icon: "⛽" },
    { text: "Setting the stage for success...", icon: "🎭" },
    { text: "Dialing in your focus...", icon: "🎚️" },
    { text: "Loading today's game plan...", icon: "🏈" },
    { text: "Waking up the network...", icon: "🌐" },
    { text: "Preparing something great...", icon: "🎁" },
    { text: "Syncing mind and machine...", icon: "🤖" },
    { text: "Your day is about to level up...", icon: "🎮" },
    { text: "Deploying your daily toolkit...", icon: "🧰" },
    { text: "Reading the tea leaves...", icon: "🍵" },
    { text: "Initializing beast mode...", icon: "🐉" },
    { text: "Weaving your productivity web...", icon: "🕸️" },
    { text: "Firing up the flux capacitor...", icon: "⚡" },
    { text: "Counting down to liftoff...", icon: "🚀" },
    { text: "Scouting the terrain ahead...", icon: "🏕️" },
    { text: "Syncing your satellite feed...", icon: "🛰️" },
    { text: "Activating your sixth sense...", icon: "👁️" },
    { text: "Booting up mission control...", icon: "🖥️" },
    { text: "Collecting today's treasures...", icon: "💰" },
    { text: "Mixing your morning potion...", icon: "🧪" },
    { text: "Rolling out the red carpet...", icon: "🎪" },
    { text: "Untangling the threads...", icon: "🧵" },
    { text: "Plugging into the matrix...", icon: "🔌" },
    { text: "Dusting off the crystal ball...", icon: "✨" },
    { text: "Loading your cheat codes...", icon: "🕹️" },
    { text: "Assembling your dream team...", icon: "🤝" },
    { text: "Plotting your next move...", icon: "♟️" },
    { text: "Unfurling the treasure map...", icon: "🗺️" },
    { text: "Shaking the magic 8-ball...", icon: "🎱" },
    { text: "Tuning your radar...", icon: "📡" },
    { text: "Stoking the creative fire...", icon: "🔥" },
    { text: "Polishing the brass...", icon: "🔔" },
    { text: "Raising the sails...", icon: "⛵" },
    { text: "Clearing the runway...", icon: "🛬" },
    { text: "Decrypting your priorities...", icon: "🔐" },
    { text: "Opening the vault...", icon: "🏦" },
    { text: "Shuffling the deck in your favor...", icon: "🃏" },
    { text: "Setting coordinates...", icon: "🧭" },
    { text: "Amplifying your signal...", icon: "📢" },
    { text: "Forging your daily armor...", icon: "🛡️" },
    { text: "Painting today's canvas...", icon: "🎨" },
    { text: "Winding up the clockwork...", icon: "⏰" },
    { text: "Channeling good vibes...", icon: "🌈" },
    { text: "Mining the data goldmine...", icon: "⛏️" },
    { text: "Stretching before the sprint...", icon: "🏃" },
    { text: "Flipping the power switch...", icon: "🔋" },
    { text: "Composing your symphony...", icon: "🎻" },
    { text: "Building momentum...", icon: "🌪️" },
    { text: "Planting seeds of productivity...", icon: "🌱" },
    { text: "Wiring up the neurons...", icon: "⚙️" },
    { text: "Stacking the deck in your favor...", icon: "🎰" },
    { text: "Focusing the telescope...", icon: "🔭" },
    { text: "Loading the knowledge base...", icon: "📖" },
    { text: "Engaging hyperdrive...", icon: "🌠" },
    { text: "Cracking the code...", icon: "🧬" },
    { text: "Preparing your runway...", icon: "✈️" },
    { text: "One moment of magic coming up...", icon: "🪄" },
  ];
  const splashPick = splashQuotes[Math.floor(Math.random() * splashQuotes.length)];
  const splashPick2 = splashQuotes.filter(q => q !== splashPick)[Math.floor(Math.random() * (splashQuotes.length - 1))];

  const isDarkSplash = document.body.classList.contains('dark-mode');
  const splash = document.createElement('div');
  splash.id = 'oracle-splash-loader';
  splash.style.cssText = `position:fixed;top:0;left:0;right:0;bottom:0;z-index:99999;display:flex;flex-direction:column;align-items:center;justify-content:center;background:${isDarkSplash ? '#0f0f1a' : 'linear-gradient(135deg, #667eea 0%, #764ba2 100%)'};transition:opacity 0.5s ease-out;`;
  splash.innerHTML = `
    <style>
      @keyframes oracleSplashPulse { 0%,100% { transform:scale(1); opacity:0.9; } 50% { transform:scale(1.1); opacity:1; } }
      @keyframes oracleSplashSpin { to { transform:rotate(360deg); } }
      @keyframes oracleSplashFadeUp { 0% { opacity:0; transform:translateY(10px); } 100% { opacity:1; transform:translateY(0); } }
      @keyframes oracleSplashQuoteSwap { 0% { opacity:1; transform:translateY(0); } 45% { opacity:0; transform:translateY(-12px); } 55% { opacity:0; transform:translateY(12px); } 100% { opacity:1; transform:translateY(0); } }
      #oracle-splash-loader .splash-icon { font-size:48px; animation:oracleSplashPulse 2s ease-in-out infinite; margin-bottom:24px; }
      #oracle-splash-loader .splash-spinner { width:40px; height:40px; border:3px solid rgba(255,255,255,0.2); border-top-color:white; border-radius:50%; animation:oracleSplashSpin 0.9s linear infinite; margin-bottom:28px; }
      #oracle-splash-loader .splash-quote { font-family:'Lato',sans-serif; font-size:18px; font-weight:500; color:rgba(255,255,255,0.95); text-align:center; max-width:340px; line-height:1.5; animation:oracleSplashFadeUp 0.6s ease-out; }
      #oracle-splash-loader .splash-sub { font-family:'Lato',sans-serif; font-size:13px; color:rgba(255,255,255,0.5); margin-top:12px; letter-spacing:0.5px; }
    </style>
    <div class="splash-icon">${splashPick.icon}</div>
    <div class="splash-spinner"></div>
    <div class="splash-quote" id="splashQuoteText">${splashPick.text}</div>
    <div class="splash-sub">Oracle is getting ready</div>
  `;
  document.body.appendChild(splash);

  // Swap quote halfway through the delay
  const splashQuoteSwapTimer = setTimeout(() => {
    const quoteEl = document.getElementById('splashQuoteText');
    if (quoteEl) {
      quoteEl.style.animation = 'oracleSplashQuoteSwap 0.6s ease-in-out';
      setTimeout(() => {
        quoteEl.textContent = splashPick2.text;
        const iconEl = splash.querySelector('.splash-icon');
        if (iconEl) iconEl.textContent = splashPick2.icon;
      }, 300);
    }
  }, INITIAL_LOAD_DELAY / 2);

  function removeSplash() {
    clearTimeout(splashQuoteSwapTimer);
    const el = document.getElementById('oracle-splash-loader');
    if (!el) return;
    el.style.opacity = '0';
    setTimeout(() => el.remove(), 500);
  }

  function startInitialLoad() {
    removeSplash();
    if (isThreeCol) {
      loadTodos('starred');
      loadFYI();
      loadDailyFeed();
    } else {
      loadTodos('starred');
    }
  }

  // Only load data if the tab stays visible for 3 seconds
  if (!document.hidden) {
    initialLoadTimer = setTimeout(startInitialLoad, INITIAL_LOAD_DELAY);
  }

  // Cancel the delayed load if the tab becomes hidden before 3s
  // Start the load when the tab becomes visible again
  const visibilityHandler = () => {
    if (document.hidden) {
      if (initialLoadTimer) { clearTimeout(initialLoadTimer); initialLoadTimer = null; }
    } else if (initialLoadTimer === null && !window._oracleInitialLoadDone) {
      initialLoadTimer = setTimeout(startInitialLoad, INITIAL_LOAD_DELAY);
    }
  };
  document.addEventListener('visibilitychange', visibilityHandler);

  // Mark initial load as done after first successful load so subsequent visibility changes don't re-trigger
  const origStartInitialLoad = startInitialLoad;
  startInitialLoad = function() {
    window._oracleInitialLoadDone = true;
    document.removeEventListener('visibilitychange', visibilityHandler);
    origStartInitialLoad();
  };

  window.loadTodos = loadTodos;
  window.loadFYI = loadFYI;
  window.loadTodosAnimated = loadTodosAnimated;
  window.loadFYIAnimated = loadFYIAnimated;
  window.loadDailyFeed = loadDailyFeed;
  window.showFeedOracleSlider = showFeedOracleSlider;
  window.showFeedDiscussionSlider = showFeedDiscussionSlider;
  window.updateDailyFeedCount = updateDailyFeedCount;
  window.displayTodos = displayTodos;
  window.displayTodosAnimated = displayTodosAnimated;
  window.displayFYI = displayFYI;
  window.loadBookmarks = loadBookmarks;
  window.loadNotes = loadNotes;
  window.loadDocuments = loadDocuments;
  window.showTranscriptSlider = showTranscriptSlider;

  // openTaskFromChat — called from assistant task chips
  // If the task isn't in memory, inject a minimal stub and open the slider immediately
  window.openTaskFromChat = async function(taskId) {
    const existing = allTodos.find(t => t.id == taskId) || allFyiItems.find(t => t.id == taskId) || allCalendarItems.find(t => t.id == taskId) || allCompletedTasks.find(t => t.id == taskId);
    if (existing) {
      showTranscriptSlider(taskId);
      return;
    }
    // Task not in memory — inject a minimal stub so slider opens instantly
    const minimalTodo = {
      id: parseInt(taskId),
      task_title: `Task #${taskId}`,
      message_link: '',
      _isStub: true
    };
    allTodos.push(minimalTodo);
    showTranscriptSlider(taskId);

    // Fetch full details in background and update the stub
    try {
      const resp = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'fetch_task_details',
          todo_id: taskId,
          timestamp: new Date().toISOString(),
          source: 'oracle-chat-task-chip'
        }))
      });
      if (resp.ok) {
        const text = await resp.text();
        if (text && text.trim()) {
          try {
            const data = JSON.parse(text);
            const rd = Array.isArray(data) ? data[0] : data;
            // Update the stub in-place
            Object.assign(minimalTodo, {
              task_title: rd?.task_title || rd?.title || minimalTodo.task_title,
              message_link: rd?.message_link || '',
              task_name: rd?.task_name || '',
              _isStub: false
            });
          } catch(e) {}
        }
      }
    } catch (err) {
      console.error('Failed to fetch task details:', err);
    }
  };
  window.updateTodoField = updateTodoField;
  window.showNoteForm = showNoteForm;
  window.hideNoteForm = hideNoteForm;
  window.showNoteViewer = window.OracleNotes.showNoteViewer;
  window.showNoteEditSlider = window.OracleNotes.showNoteEditSlider;

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
  // Normalize to string for consistent comparison
  const id = String(todoId);
  if (selectedTodoIds.has(id)) {
    selectedTodoIds.delete(id);
  } else {
    selectedTodoIds.add(id);
  }
  // Also remove numeric version if present (cleanup mixed types)
  const numId = parseInt(id);
  if (!isNaN(numId) && selectedTodoIds.has(numId)) {
    selectedTodoIds.delete(numId);
    if (!selectedTodoIds.has(id)) selectedTodoIds.add(id);
  }
  if (selectedTodoIds.size === 0) isMultiSelectMode = false;
  updateSelectionUI();
}

function clearMultiSelection() {
  selectedTodoIds.clear();
  isMultiSelectMode = false;
  updateSelectionUI();
}

function updateSelectionUI() {
  // Update todo item visual state — only highlight the .multi-selected class
  // on items that are actually in the same container as where selection started.
  // This prevents cross-section highlighting when IDs exist in both Action and FYI.
  const fyiContainer = document.querySelector('.fyi-container, #fyiColumn');
  const actionContainer = document.querySelector('#actionColumn');

  // Determine which container the selection started from (based on first selected item found in DOM)
  let selectionContainer = null;
  if (!selectionContainer) {
    for (const id of selectedTodoIds) {
      const el = document.querySelector(`.todo-item.multi-selected[data-todo-id="${id}"]`);
      if (el) {
        if (fyiContainer && fyiContainer.contains(el)) { selectionContainer = 'fyi'; break; }
        if (actionContainer && actionContainer.contains(el)) { selectionContainer = 'action'; break; }
      }
    }
  }
  // If no existing selection, detect from newly added items
  if (!selectionContainer && selectedTodoIds.size > 0) {
    for (const id of selectedTodoIds) {
      // Check FYI first
      if (fyiContainer && fyiContainer.querySelector(`.todo-item[data-todo-id="${id}"]`)) { selectionContainer = 'fyi'; break; }
      if (actionContainer && actionContainer.querySelector(`.todo-item[data-todo-id="${id}"]`)) { selectionContainer = 'action'; break; }
    }
  }

  document.querySelectorAll('.todo-item').forEach(item => {
    const todoId = String(item.dataset.todoId);
    if (selectedTodoIds.has(todoId)) {
      // Only highlight if item is in the same section as the selection
      const inFyi = fyiContainer && fyiContainer.contains(item);
      const inAction = actionContainer && actionContainer.contains(item);
      if ((selectionContainer === 'fyi' && inFyi) || (selectionContainer === 'action' && inAction) || !selectionContainer) {
        item.classList.add('multi-selected');
      } else {
        item.classList.remove('multi-selected');
      }
    } else {
      item.classList.remove('multi-selected');
    }
  });

  // Update meeting item visual state
  document.querySelectorAll('.meeting-item').forEach(item => {
    const meetingId = String(item.dataset.meetingId);
    if (selectedTodoIds.has(meetingId)) {
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

  // Update per-column selection count badges
  const actionSelBadge = document.getElementById('actionSelectionCount');
  const fyiSelBadge = document.getElementById('fyiSelectionCount');
  let actionSelCount = 0, fyiSelCount = 0;
  if (selectedTodoIds.size > 0) {
    for (const id of selectedTodoIds) {
      if (actionContainer && actionContainer.querySelector(`.todo-item[data-todo-id="${id}"]`)) actionSelCount++;
      else if (fyiContainer && fyiContainer.querySelector(`.todo-item[data-todo-id="${id}"]`)) fyiSelCount++;
    }
  }
  if (actionSelBadge) {
    if (actionSelCount > 0) { actionSelBadge.textContent = `Mark ${actionSelCount} item${actionSelCount > 1 ? 's' : ''} as done`; actionSelBadge.style.display = 'inline-flex'; actionSelBadge.onclick = () => handleBulkMarkDone(); }
    else { actionSelBadge.style.display = 'none'; actionSelBadge.onclick = null; }
  }
  if (fyiSelBadge) {
    if (fyiSelCount > 0) { fyiSelBadge.textContent = `Mark ${fyiSelCount} item${fyiSelCount > 1 ? 's' : ''} as done`; fyiSelBadge.style.display = 'inline-flex'; fyiSelBadge.onclick = () => handleBulkMarkDone(); }
    else { fyiSelBadge.style.display = 'none'; fyiSelBadge.onclick = null; }
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

  // Determine which section each selected item lives in by checking its DOM container
  const fyiIds = new Set();
  const actionIds = new Set();
  const otherIds = new Set();
  const fyiContainer = document.querySelector('.fyi-container, #fyiColumn');
  const actionContainer = document.querySelector('#actionColumn');

  selectedIds.forEach(id => {
    const strId = String(id);
    const el = document.querySelector(`.todo-item.multi-selected[data-todo-id="${strId}"]`);
    if (el && fyiContainer && fyiContainer.contains(el)) {
      fyiIds.add(strId);
    } else if (el && actionContainer && actionContainer.contains(el)) {
      actionIds.add(strId);
    } else {
      // Fallback: check array membership
      if (allFyiItems.some(t => String(t.id) === strId)) fyiIds.add(strId);
      else if (allTodos.some(t => String(t.id) === strId)) actionIds.add(strId);
      else otherIds.add(strId);
    }
  });

  // Immediately animate and remove selected items from UI
  selectedIds.forEach(todoId => {
    const strId = String(todoId);
    let todoItem = null;
    // Only look in the container where this ID was selected
    if (fyiIds.has(strId) && fyiContainer) {
      todoItem = fyiContainer.querySelector(`.todo-item[data-todo-id="${todoId}"]`);
    } else if (actionIds.has(strId) && actionContainer) {
      todoItem = actionContainer.querySelector(`.todo-item[data-todo-id="${todoId}"]`);
    }
    // No global fallback — only remove from the section where item was selected
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

  // Update local arrays immediately (optimistic update) — only remove from the section the item was selected from
  if (actionIds.size > 0) {
    allTodos = allTodos.filter(t => !actionIds.has(String(t.id)));
  }
  if (fyiIds.size > 0) {
    allFyiItems = allFyiItems.filter(t => !fyiIds.has(String(t.id)));
  }

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
  clearKeyboardSelection();
  currentKeyboardColumn = null;

  // Show success toast
  showBulkSuccessToast(selectedTodos.length);

  // Update counts
  updateTabCounts();

  // Clean up empty task groups after animation
  setTimeout(() => {
    document.querySelectorAll('.task-group').forEach(group => {
      const remaining = group.querySelectorAll('.todo-item:not(.completing)');
      if (remaining.length === 0) {
        group.remove();
      } else {
        const countEl = group.querySelector('.task-group-count');
        if (countEl) {
          countEl.textContent = `${remaining.length} item${remaining.length > 1 ? 's' : ''}`;
        }
      }
    });
    updateTabCounts();
  }, 450);

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
  showToastNotification(`✓ ${count} task${count > 1 ? 's' : ''} marked as done`);
}

document.addEventListener('DOMContentLoaded', async () => {
  // Clear extension badge on fresh page load (no pending updates yet)
  updateExtensionBadge();
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
// V38: Ably removed — using native n8n webhook streaming

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
  newMsgToggle.addEventListener('click', () => { window.OracleNewMessage.showNewMessageSlider({ mode: 'col3' }); });
}
function showNewMessageSlider() { window.OracleNewMessage.showNewMessageSlider({ mode: 'col3' }); }

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

  if (message.action === 'openTaskTranscript') {
    const taskId = message.taskId;
    if (taskId && typeof window.openTaskFromChat === 'function') {
      window.openTaskFromChat(taskId);
    }
    sendResponse({ received: true });
    return true;
  }

  // Handle extracted_query from transcript flow via Ably
  if (message.action === 'oracleSystemMessage' && message.data?.type === 'extracted_query') {
    console.log('🔮 Oracle extracted query received:', message.data.text);
    // Find the open chat slider with reading state and resolve it
    const chatOverlay = document.querySelector('.chat-slider-overlay');
    if (chatOverlay && chatOverlay._oracleResolveQuery) {
      chatOverlay._oracleResolveQuery(message.data.text);
    } else if (!isChatSliderOpen) {
      // No slider open — open one and auto-send the extracted query
      window.OracleAssistant.showChatSlider({
        mode: 'fullscreen',
        prefillMessage: message.data.text,
        onClose: () => { isChatSliderOpen = false; }
      });
      isChatSliderOpen = true;
    }
    sendResponse({ received: true });
    return true;
  }

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

      // Refresh Meetings accordion immediately only if user is NOT active on the tab
      const newMeetings = activeActionTodos.filter(t => isMeetingLink(t.message_link));
      if (newMeetings.length > 0) {
        if (document.hidden || !document.hasFocus()) {
          refreshMeetingsAccordionOnly(newMeetings, activeActionTodos);
        }
        // If user is active, the "X new updates" banner will prompt them to refresh
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

      // Refresh Documents accordion immediately only if user is NOT active on the tab
      const newDocs = newFyiTodos.filter(t => isDriveLink(t.message_link));
      if (newDocs.length > 0) {
        if (document.hidden || !document.hasFocus()) {
          refreshDocumentsAccordionOnly(newDocs, newFyiTodos);
        }
        // If user is active, the "X new updates" banner will prompt them to refresh
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
  updateExtensionBadge(); // Sync Chrome extension badge
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

    // Refresh Daily Feed tab as well (new feed items may have arrived)
    if (typeof window.loadDailyFeed === 'function') window.loadDailyFeed();

    // Cross-deduplicate: if a task moved from FYI→Action or Action→FYI, remove it from the old list.
    // Always run this regardless of whether pending data existed for each tab — a task may have
    // moved between tabs without triggering a "new/updated" signal on the source side.
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

    // Always re-render both sides so that moved tasks disappear from the old tab,
    // even if that tab had no pending data of its own.
    const fyiContainer = document.querySelector('.fyi-container');
    if (fyiRemoved > 0 || pendingFyiData) {
      if (fyiContainer && typeof window.displayFYI === 'function') {
        window.displayFYI(allFyiItems, fyiContainer, []);
      }
    }
    if (actionRemoved > 0 || pendingActionData) {
      if (typeof window.displayTodos === 'function') {
        window.displayTodos(allTodos, currentTodoFilter || 'starred', []);
      }
    }
    if (fyiRemoved > 0 || actionRemoved > 0) {
      console.log(`🔀 Cross-dedup: removed ${fyiRemoved} from FYI (moved to Action), ${actionRemoved} from Action (moved to FYI)`);
    }
    if (typeof window.updateTabCounts === 'function') window.updateTabCounts();

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

    // Mark updated tasks as unread (skip is_latest_from_self)
    pendingActionUpdates.forEach(id => {
      const task = newTodos.find(t => t.id == id);
      if (task && task.is_latest_from_self === true) return;
      readTaskIds.delete(String(id));
    });

    // Force-read any is_latest_from_self items
    newTodos.forEach(t => {
      if (t.is_latest_from_self === true) readTaskIds.add(String(t.id));
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
    updateExtensionBadge();

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

    // Mark updated tasks as unread (skip is_latest_from_self)
    pendingFyiUpdates.forEach(id => {
      const task = newFyiItems.find(t => t.id == id);
      if (task && task.is_latest_from_self === true) return;
      readTaskIds.delete(String(id));
    });

    // Force-read any is_latest_from_self items
    newFyiItems.forEach(t => {
      if (t.is_latest_from_self === true) readTaskIds.add(String(t.id));
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
    updateExtensionBadge();

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
let profileMonitors = []; // Array of monitors from Parallel API
let originalMonitors = []; // Store original monitors for comparison

// escapeHtml already available from shared component alias (line 54)

let _oracleToastEl = null;
let _oracleToastTimer = null;
function showToastNotification(message) {
  // Reuse a single toast element — prevents stacking and jitter
  if (_oracleToastTimer) clearTimeout(_oracleToastTimer);

  if (!_oracleToastEl) {
    _oracleToastEl = document.createElement('div');
    _oracleToastEl.style.cssText = `
      position: fixed;
      top: 20px;
      left: 50%;
      transform: translateX(-50%);
      background: linear-gradient(45deg, #667eea, #764ba2);
      color: white;
      padding: 12px 24px;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 600;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
      z-index: 10001;
      white-space: nowrap;
      transition: opacity 0.3s ease-out;
      opacity: 0;
      pointer-events: none;
    `;
    document.body.appendChild(_oracleToastEl);
  }
  _oracleToastEl.textContent = message;
  _oracleToastEl.style.opacity = '1';

  _oracleToastTimer = setTimeout(() => {
    _oracleToastEl.style.opacity = '0';
    _oracleToastTimer = null;
  }, 2000);
}

// --- Change Detection ---
function getCurrentFormData() {
  const extractIds = (arr) => {
    if (!arr) return [];
    if (typeof arr === 'string') { try { arr = JSON.parse(arr); } catch { return []; } }
    if (!Array.isArray(arr)) return [];
    return arr.map(ch => typeof ch === 'object' ? ch.id : ch);
  };
  return {
    email_ID: document.getElementById('profileEmail')?.value || '',
    password: document.getElementById('profilePassword')?.value || '',
    fr_api_key: document.getElementById('profileFrApiKey')?.value || '',
    fd_l2_api_key: document.getElementById('profileFdApiKey')?.value || '',
    slack_user_token: document.getElementById('profileSlackToken')?.value || '',
    role_description: document.getElementById('profileRoleDesc')?.value || '',
    muted_channels: extractIds(userProfileData?.muted_channels),
    monitor_channels: extractIds(userProfileData?.monitor_channels),
    allowed_public_channels: extractIds(userProfileData?.allowed_public_channels),
    blocked_email_participants: userProfileData?.blocked_email_participants || []
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

  // Check monitor channels
  const origMonitor = originalProfileData.monitor_channels || [];
  const currMonitor = current.monitor_channels || [];
  if (JSON.stringify(origMonitor) !== JSON.stringify(currMonitor)) return true;

  // Check allowed public channels
  const origAllowed = originalProfileData.allowed_public_channels || [];
  const currAllowed = current.allowed_public_channels || [];
  if (JSON.stringify(origAllowed) !== JSON.stringify(currAllowed)) return true;

  // Check blocked email participants
  const origBlocked = originalProfileData.blocked_email_participants || [];
  const currBlocked = current.blocked_email_participants || [];
  if (JSON.stringify(origBlocked) !== JSON.stringify(currBlocked)) return true;

  // Check tags
  const cleanCurrentTags = profileTags.filter(t => t.name.trim() !== '');
  if (JSON.stringify(cleanCurrentTags) !== JSON.stringify(originalTags)) return true;

  // Check monitors
  const cleanMonitors = profileMonitors.filter(m => m.query.trim() !== '').map(m => ({ monitor_id: m.monitor_id, query: m.query, isNew: m.isNew }));
  const cleanOrigMonitors = originalMonitors.filter(m => m.query.trim() !== '').map(m => ({ monitor_id: m.monitor_id, query: m.query, isNew: m.isNew }));
  if (JSON.stringify(cleanMonitors) !== JSON.stringify(cleanOrigMonitors)) return true;

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

  // Bind textarea changes + auto-resize
  tbody.querySelectorAll('textarea').forEach(textarea => {
    // Auto-resize on load (use requestAnimationFrame for accurate scrollHeight)
    requestAnimationFrame(() => {
      textarea.style.height = 'auto';
      textarea.style.height = textarea.scrollHeight + 'px';
    });
    textarea.addEventListener('input', (e) => {
      // Auto-resize
      e.target.style.height = 'auto';
      e.target.style.height = e.target.scrollHeight + 'px';
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

// ============================================
// MONITORS TABLE
// ============================================
function renderMonitorsTable() {
  const tbody = document.getElementById('monitorsTableBody');
  if (!tbody) return;

  tbody.innerHTML = profileMonitors.map((m, i) => `
    <div class="monitor-row ${m.isNew ? 'monitor-row-new' : ''}" data-index="${i}">
      <textarea placeholder="e.g. Extract recent news on pricing changes from Intercom Fin AI…" data-field="query" rows="1">${escapeHtml(m.query)}</textarea>
      <button class="monitor-row-remove" data-index="${i}" title="Delete monitor">×</button>
    </div>
  `).join('');

  // Bind textarea changes — only update local state, save happens via Save button
  tbody.querySelectorAll('textarea').forEach(el => {
    // Auto-resize on load
    requestAnimationFrame(() => {
      el.style.height = 'auto';
      el.style.height = el.scrollHeight + 'px';
    });
    el.addEventListener('input', (e) => {
      // Auto-resize
      e.target.style.height = 'auto';
      e.target.style.height = e.target.scrollHeight + 'px';
      const row = e.target.closest('.monitor-row');
      const idx = parseInt(row.dataset.index);
      const field = e.target.dataset.field;
      const monitor = profileMonitors[idx];
      if (!monitor) return;
      monitor[field] = e.target.value;
      updateSaveButtonVisibility();
    });
  });

  // Bind remove buttons
  tbody.querySelectorAll('.monitor-row-remove').forEach(btn => {
    btn.addEventListener('click', async (e) => {
      const idx = parseInt(btn.dataset.index);
      const monitor = profileMonitors[idx];
      if (!monitor) return;

      if (!confirm(`Delete monitor:\n"${monitor.query.substring(0, 80)}..."?`)) return;

      // Remove locally — actual API deletion happens on Save
      profileMonitors.splice(idx, 1);
      renderMonitorsTable();
      updateSaveButtonVisibility();
    });
  });
}

function addMonitorRow() {
  profileMonitors.push({
    monitor_id: null,
    query: '',
    status: 'active',
    cadence: 'daily',
    isNew: true
  });
  renderMonitorsTable();
  // Focus the new query textarea
  const tbody = document.getElementById('monitorsTableBody');
  const lastRow = tbody.lastElementChild;
  if (lastRow) lastRow.querySelector('textarea[data-field="query"]').focus();
}

async function createNewMonitor(index) {
  const monitor = profileMonitors[index];
  if (!monitor || !monitor.isNew || !monitor.query.trim()) return;

  try {
    const response = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'add_monitor',
        query: monitor.query,
        cadence: monitor.cadence,
        timestamp: new Date().toISOString()
      }))
    });
    const data = await response.json();
    // Update local state with returned monitor_id
    monitor.isNew = false;
    monitor.monitor_id = data.monitor_id || data.id;
    renderMonitorsTable();
    showToastNotification('Monitor created!');
  } catch (err) {
    console.error('Error creating monitor:', err);
    showToastNotification('Failed to create monitor');
  }
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

      // Parse blocked_email_participants from JSON string if needed
      if (typeof userProfileData.blocked_email_participants === 'string') {
        try {
          userProfileData.blocked_email_participants = JSON.parse(userProfileData.blocked_email_participants);
        } catch (e) {
          userProfileData.blocked_email_participants = [];
        }
      }
      if (!Array.isArray(userProfileData.blocked_email_participants)) {
        userProfileData.blocked_email_participants = [];
      }

      // Load tags from profile data (include IDs for tracking changes)
      if (Array.isArray(userProfileData.tags)) {
        profileTags = userProfileData.tags.map(t => ({
          id: t.id || null,
          name: t.tag_name || t.name || '',
          description: t.tag_description || t.description || ''
        }));
      } else {
        profileTags = [];
      }
      // Store original data for change detection
      const extractIdsOrig = (arr) => {
        if (!arr) return [];
        if (typeof arr === 'string') { try { arr = JSON.parse(arr); } catch { return []; } }
        if (!Array.isArray(arr)) return [];
        return arr.map(ch => typeof ch === 'object' ? ch.id : ch);
      };
      originalProfileData = {
        email_ID: userProfileData.email_ID || '',
        password: userProfileData.password || '',
        fr_api_key: userProfileData.fr_api_key || '',
        fd_l2_api_key: userProfileData.fd_l2_api_key || '',
        slack_user_token: userProfileData.slack_user_token || '',
        role_description: userProfileData.role_description || '',
        muted_channels: extractIdsOrig(userProfileData.muted_channels),
        monitor_channels: extractIdsOrig(userProfileData.monitor_channels),
        allowed_public_channels: extractIdsOrig(userProfileData.allowed_public_channels),
        blocked_email_participants: [...(userProfileData.blocked_email_participants || [])]
      };
      // Store original tags with IDs for change tracking
      originalTags = profileTags.filter(t => t.name.trim() !== '').map(t => ({ ...t }));

      // Load monitors from profile data
      if (Array.isArray(userProfileData.monitors)) {
        profileMonitors = userProfileData.monitors.map(m => ({
          monitor_id: m.monitor_id,
          query: m.query || '',
          status: m.status || 'active',
          cadence: m.cadence || 'daily',
          created_at: m.created_at,
          last_run_at: m.last_run_at,
          isNew: false
        }));
      } else {
        profileMonitors = [];
      }
      // Store original monitors for change tracking
      originalMonitors = profileMonitors.map(m => ({ ...m }));

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

  // Display monitor channels as blue tags
  let monChannels = userProfileData.monitor_channels || [];
  if (typeof monChannels === 'string') { try { monChannels = JSON.parse(monChannels); } catch { monChannels = []; } }
  displayMonitorChannelTags(monChannels);

  // Display allowed public channels as tags
  displayAllowedChannelTags(userProfileData.allowed_public_channels || []);

  // Display blocked email participants as tags
  displayBlockedEmailTags(userProfileData.blocked_email_participants || []);

  // Render tags table
  renderTagsTable();

  // Render monitors table
  renderMonitorsTable();
}

function displayChannelTags(channels) {
  const container = document.getElementById('profileChannelTags');
  const channelsArray = Array.isArray(channels) ? channels : (channels ? [channels] : []);

  container.innerHTML = channelsArray.map(ch => {
    const id = typeof ch === 'object' ? ch.id : ch;
    const name = typeof ch === 'object' ? (ch.name || ch.id) : ch;
    return `<div class="profile-tag">
      ${escapeHtml(name)}
      <span class="profile-tag-remove" data-channel-id="${escapeHtml(id)}">×</span>
    </div>`;
  }).join('');

  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      removeChannel(btn.getAttribute('data-channel-id'));
    });
  });
}

function removeChannel(channelId) {
  if (!userProfileData.muted_channels) userProfileData.muted_channels = [];
  userProfileData.muted_channels = userProfileData.muted_channels.filter(ch =>
    (typeof ch === 'object' ? ch.id : ch) !== channelId
  );
  displayChannelTags(userProfileData.muted_channels);
  updateSaveButtonVisibility();
}

function displayMonitorChannelTags(channels) {
  const container = document.getElementById('profileMonitorChannelTags');
  if (!container) return;
  const channelsArray = Array.isArray(channels) ? channels : (channels ? [channels] : []);

  container.innerHTML = channelsArray.map(ch => {
    const id = typeof ch === 'object' ? ch.id : ch;
    const name = typeof ch === 'object' ? (ch.name || ch.id) : ch;
    return `<div class="profile-tag" style="background: linear-gradient(45deg, #f39c12, #e67e22); border-color: transparent; color: #1a1a1a;">
      ${escapeHtml(name)}
      <span class="profile-tag-remove" data-channel-id="${escapeHtml(id)}" style="color:#1a1a1a;">×</span>
    </div>`;
  }).join('');

  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      removeMonitorChannel(btn.getAttribute('data-channel-id'));
    });
  });
}

function removeMonitorChannel(channelId) {
  if (!userProfileData.monitor_channels) userProfileData.monitor_channels = [];
  let channels = userProfileData.monitor_channels;
  if (typeof channels === 'string') { try { channels = JSON.parse(channels); } catch { channels = []; } }
  if (!Array.isArray(channels)) channels = [];
  userProfileData.monitor_channels = channels.filter(ch =>
    (typeof ch === 'object' ? ch.id : ch) !== channelId
  );
  displayMonitorChannelTags(userProfileData.monitor_channels);
  updateSaveButtonVisibility();
}

function displayAllowedChannelTags(channels) {
  const container = document.getElementById('profileAllowedChannelTags');
  if (!container) return;
  const channelsArray = Array.isArray(channels) ? channels : (channels ? [channels] : []);

  container.innerHTML = channelsArray.map(ch => {
    const id = typeof ch === 'object' ? ch.id : ch;
    const name = typeof ch === 'object' ? (ch.name || ch.id) : ch;
    return `<div class="profile-tag" style="background: linear-gradient(45deg, #27ae60, #2ecc71); border-color: transparent;">
      ${escapeHtml(name)}
      <span class="profile-tag-remove" data-channel-id="${escapeHtml(id)}">×</span>
    </div>`;
  }).join('');

  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      removeAllowedChannel(btn.getAttribute('data-channel-id'));
    });
  });
}

function removeAllowedChannel(channelId) {
  if (!userProfileData.allowed_public_channels) userProfileData.allowed_public_channels = [];
  userProfileData.allowed_public_channels = userProfileData.allowed_public_channels.filter(ch =>
    (typeof ch === 'object' ? ch.id : ch) !== channelId
  );
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

  // Helper: extract plain ID strings from channel arrays (which may be {id,name} objects or strings)
  const extractIds = (arr) => (arr || []).map(ch => typeof ch === 'object' ? ch.id : ch);

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
    muted_channels: extractIds(userProfileData.muted_channels),
    monitor_channels: extractIds(userProfileData.monitor_channels),
    allowed_public_channels: extractIds(userProfileData.allowed_public_channels),
    blocked_email_participants: userProfileData.blocked_email_participants || [],
    tags_added: tagsAdded,
    tags_modified: tagsModified,
    tags_removed: tagsRemoved,
    timestamp: new Date().toISOString()
  };

  // Calculate monitor changes
  const monitorsToCreate = profileMonitors.filter(m => m.isNew && m.query.trim());
  const originalMonitorIds = new Set(originalMonitors.map(m => m.monitor_id).filter(Boolean));
  const currentMonitorIds = new Set(profileMonitors.map(m => m.monitor_id).filter(Boolean));
  const monitorsToDelete = originalMonitors.filter(m => m.monitor_id && !currentMonitorIds.has(m.monitor_id));
  const monitorsToUpdate = profileMonitors.filter(m => {
    if (m.isNew || !m.monitor_id) return false;
    const orig = originalMonitors.find(om => om.monitor_id === m.monitor_id);
    return orig && orig.query !== m.query;
  });

  console.log('Saving profile with tag changes:', { tagsAdded, tagsModified, tagsRemoved });
  console.log('Monitor changes:', { monitorsToCreate: monitorsToCreate.length, monitorsToUpdate: monitorsToUpdate.length, monitorsToDelete: monitorsToDelete.length });

  try {
    await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload(updatedData))
    });

    // Handle monitor CRUD operations
    const monitorPromises = [];

    // Create new monitors
    for (const m of monitorsToCreate) {
      monitorPromises.push(fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'add_monitor',
          query: m.query,
          cadence: m.cadence || 'daily',
          timestamp: new Date().toISOString()
        }))
      }));
    }

    // Update modified monitors
    for (const m of monitorsToUpdate) {
      monitorPromises.push(fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'update_monitor',
          monitor_id: m.monitor_id,
          query: m.query,
          cadence: m.cadence || 'daily',
          timestamp: new Date().toISOString()
        }))
      }));
    }

    // Delete removed monitors
    for (const m of monitorsToDelete) {
      monitorPromises.push(fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'delete_monitor',
          monitor_id: m.monitor_id,
          timestamp: new Date().toISOString()
        }))
      }));
    }

    if (monitorPromises.length > 0) {
      await Promise.allSettled(monitorPromises);
    }

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

// ============================================
// BLOCKED EMAIL PARTICIPANTS — Block modal + profile section
// ============================================

// ============================================
// BELL DROPDOWN — Show mute/monitor options for Slack channels
// ============================================
function showBellDropdown(e, channelId, channelName) {
  if (!channelId) return;
  e.stopPropagation();
  // Resolve channel name from DOM context if not provided
  if (!channelName || channelName === channelId) {
    const btn = e.currentTarget || e.target.closest('.slack-bell-channel-btn');
    if (btn) {
      // Try data attribute first
      channelName = btn.dataset.channelName;
      if (!channelName) {
        // Try finding participant_text or group header nearby
        const parent = btn.closest('.todo-item, .task-group-task-item, .slack-channel-accordion-item, .group-item, [data-todo-id]');
        if (parent) {
          const ptEl = parent.querySelector('.todo-participant, .slack-channel-item-name, .task-group-name');
          if (ptEl) channelName = ptEl.textContent.trim().replace(/^(Private|Public) channel\s*/i, '');
        }
        // Try the group header if inside a task group
        if (!channelName) {
          const group = btn.closest('.task-group, .slack-channel-accordion-item');
          if (group) {
            const nameEl = group.querySelector('.task-group-name, .slack-channel-item-name');
            if (nameEl) channelName = nameEl.textContent.trim();
          }
        }
      }
    }
    channelName = channelName || userProfileData?.channel_map?.[channelId] || channelId;
  }

  // Close any existing dropdown
  document.querySelectorAll('.slack-bell-dropdown').forEach(d => d.remove());

  const btn = e.currentTarget || e.target.closest('.slack-bell-channel-btn');
  if (!btn) return;

  // Position dropdown relative to the button using fixed coordinates
  const rect = btn.getBoundingClientRect();

  const dropdown = document.createElement('div');
  dropdown.className = 'slack-bell-dropdown show';
  dropdown.style.top = (rect.bottom + 4) + 'px';
  dropdown.style.left = (rect.right - 180) + 'px'; // align right edge

  dropdown.innerHTML = `
    <div class="slack-bell-dropdown-item" data-action="mute">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#e74c3c" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M13.73 21a2 2 0 0 1-3.46 0"/><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><line x1="1" y1="1" x2="23" y2="23"/></svg>
      <span style="color:#e74c3c;font-weight:500;">Stop monitoring</span>
    </div>
    <div class="slack-bell-dropdown-item" data-action="monitor">
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"/><path d="M13.73 21a2 2 0 0 1-3.46 0"/></svg>
      <span style="color:#667eea;font-weight:500;">Daily monitoring</span>
    </div>
  `;
  document.body.appendChild(dropdown);

  // Ensure dropdown doesn't go off-screen
  requestAnimationFrame(() => {
    const dRect = dropdown.getBoundingClientRect();
    if (dRect.left < 8) dropdown.style.left = '8px';
    if (dRect.bottom > window.innerHeight - 8) {
      dropdown.style.top = (rect.top - dRect.height - 4) + 'px';
    }
  });

  dropdown.querySelector('[data-action="mute"]').addEventListener('click', (ev) => {
    ev.stopPropagation();
    dropdown.remove();
    blockSlackChannel(channelId, channelName);
  });
  dropdown.querySelector('[data-action="monitor"]').addEventListener('click', (ev) => {
    ev.stopPropagation();
    dropdown.remove();
    addToMonitorChannels(channelId, channelName);
  });

  // Close on outside click
  const closeHandler = (ev) => {
    if (!dropdown.contains(ev.target) && !btn.contains(ev.target)) {
      dropdown.remove();
      document.removeEventListener('click', closeHandler, true);
    }
  };
  setTimeout(() => document.addEventListener('click', closeHandler, true), 10);

  // Close on Escape
  const escHandler = (ev) => {
    if (ev.key === 'Escape') {
      dropdown.remove();
      document.removeEventListener('keydown', escHandler);
      document.removeEventListener('click', closeHandler, true);
    }
  };
  document.addEventListener('keydown', escHandler);
}

// ============================================
// ADD TO MONITOR CHANNELS — Send webhook to n8n
// ============================================
async function addToMonitorChannels(channelId, channelName) {
  if (!channelId) return;
  channelName = channelName || userProfileData?.channel_map?.[channelId] || channelId;
  try {
    showToastNotification(`Adding ${channelName} to daily monitoring...`);
    await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'daily_slack_channel_monitoring',
        channel_id: channelId,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-newtab'
      }))
    });
    showToastNotification(`Added ${channelName} to daily monitoring`);
  } catch (err) {
    console.error('Error adding to monitor channels:', err);
    showToastNotification('Failed to add to monitor channels');
  }
}

// ============================================
// BLOCK SLACK CHANNEL — Send webhook to n8n
// ============================================
async function blockSlackChannel(channelId, passedChannelName) {
  if (!channelId) return;

  const isDark = document.body.classList.contains('dark-mode');
  const channelName = passedChannelName || userProfileData?.channel_map?.[channelId] || channelId;

  // Confirm with user
  const confirmed = await new Promise(resolve => {
    let resolved = false;
    const safeResolve = (val) => { if (resolved) return; resolved = true; overlay.remove(); resolve(val); };
    const overlay = document.createElement('div');
    overlay.className = 'block-channel-confirm-overlay';
    overlay.style.cssText = `position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100000;display:flex;align-items:center;justify-content:center;animation:fadeIn 0.2s ease-out;`;
    overlay.innerHTML = `
      <div style="background:${isDark ? '#1f2940' : 'white'};border-radius:16px;padding:24px;max-width:400px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,0.3);">
        <div style="font-weight:700;font-size:16px;color:${isDark ? '#e8e8e8' : '#2c3e50'};margin-bottom:8px;">Mute Channel</div>
        <div style="font-size:13px;color:${isDark ? '#aaa' : '#5d6d7e'};margin-bottom:20px;">Tasks from channel <strong style="color:${isDark ? '#ff8a80' : '#e74c3c'};">${escapeHtml(channelName)}</strong> will be hidden from your feed. You can unmute it from Profile settings.</div>
        <div style="display:flex;gap:10px;justify-content:flex-end;">
          <button class="block-ch-cancel" style="padding:10px 20px;background:${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:10px;color:${isDark ? '#b0b0b0' : '#5d6d7e'};font-size:14px;cursor:pointer;">Cancel</button>
          <button class="block-ch-confirm" style="padding:10px 20px;background:linear-gradient(45deg,#e74c3c,#c0392b);border:none;border-radius:10px;color:white;font-size:14px;font-weight:600;cursor:pointer;">Mute Channel</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    overlay.querySelector('.block-ch-cancel').addEventListener('click', (e) => { e.stopPropagation(); safeResolve(false); });
    overlay.querySelector('.block-ch-confirm').addEventListener('click', (e) => { e.stopPropagation(); safeResolve(true); });
    overlay.addEventListener('click', (e) => { if (e.target === overlay) safeResolve(false); });
  });

  if (!confirmed) return;

  try {
    await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(createAuthenticatedPayload({
        action: 'stop_slack_channel_monitoring',
        channel_id: channelId,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-newtab'
      }))
    });
    showToastNotification(`Muted ${channelName}`);

    // Refresh todos to hide muted channel tasks
    if (typeof fetchAllData === 'function') {
      fetchAllData();
    }
  } catch (err) {
    console.error('Error blocking Slack channel:', err);
    showToastNotification('Failed to mute channel');
  }
}

async function showBlockEmailModal(todoId) {
  const isDark = document.body.classList.contains('dark-mode');
  const todo = allTodos.find(t => String(t.id) === String(todoId))
    || allFyiItems.find(t => String(t.id) === String(todoId));
  if (!todo) { console.warn('🚫 Block modal: todo not found for id', todoId); return; }

  // Remove any existing block modals first
  document.querySelectorAll('.block-email-modal-overlay').forEach(m => m.remove());

  // Show loading overlay
  const overlay = document.createElement('div');
  overlay.className = 'block-email-modal-overlay';
  overlay.style.cssText = `position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100000;display:flex;align-items:center;justify-content:center;animation:fadeIn 0.2s ease-out;`;
  overlay.innerHTML = `<div style="background:${isDark ? '#1f2940' : 'white'};border-radius:16px;padding:24px;max-width:480px;width:90%;box-shadow:0 20px 60px rgba(0,0,0,0.3);"><div style="text-align:center;padding:20px;"><div class="spinner" style="margin:0 auto 12px;width:32px;height:32px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div><div style="font-size:14px;color:${isDark ? '#888' : '#7f8c8d'};">Fetching thread participants...</div></div></div>`;
  document.body.appendChild(overlay);

  try {
    const isDriveTask = todo.message_link && todo.message_link.includes('drive.google.com');
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

    const responseText = await response.text();
    let data = {};
    if (responseText && responseText.trim()) {
      try { data = JSON.parse(responseText); } catch { data = {}; }
    }
    const responseData = Array.isArray(data) ? data[0] : data;
    console.log('🚫 Block modal response:', responseData);
    const participants = responseData?.participants || [];
    const toParticipants = responseData?.to_participants || [];
    const ccParticipants = responseData?.cc_participants || [];
    console.log('🚫 Participants:', participants.length, 'To:', toParticipants.length, 'CC:', ccParticipants.length);

    // Merge all participants, dedupe by email
    const allParticipants = [];
    const seen = new Set();
    [...participants, ...toParticipants, ...ccParticipants].forEach(p => {
      const email = p.email || p.user_email_ID || '';
      if (email && !seen.has(email.toLowerCase())) {
        seen.add(email.toLowerCase());
        allParticipants.push({ name: p.name || p.Full_Name || email.split('@')[0], email });
      }
    });
    console.log('🚫 All unique participants:', allParticipants.length);

    // Also load current user email to exclude self
    const userEmail = (window.Oracle?.state?.userData?.email || '').toLowerCase();
    console.log('🚫 Current user email:', userEmail);

    // Filter out self
    const blockCandidates = allParticipants.filter(p => p.email.toLowerCase() !== userEmail);

    if (!blockCandidates.length) {
      overlay.remove();
      showToastNotification('No participants found to block');
      return;
    }

    // Render modal with participant chips
    let selectedForBlock = [...blockCandidates];

    const modal = overlay.querySelector('div');
    modal.innerHTML = `
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
        <div>
          <div style="font-weight:700;font-size:16px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Block Email Participants</div>
          <div style="font-size:12px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:4px;">Remove participants you don't want to block</div>
        </div>
        <button class="block-modal-close" style="background:none;border:none;color:${isDark ? '#888' : '#95a5a6'};font-size:22px;cursor:pointer;width:32px;height:32px;display:flex;align-items:center;justify-content:center;">×</button>
      </div>
      <div class="block-participants-list" style="display:flex;flex-wrap:wrap;gap:8px;margin-bottom:20px;max-height:200px;overflow-y:auto;padding:12px;background:${isDark ? 'rgba(255,255,255,0.03)' : 'rgba(102,126,234,0.03)'};border-radius:10px;border:1px solid ${isDark ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};"></div>
      <div style="display:flex;gap:10px;justify-content:flex-end;">
        <button class="block-modal-cancel" style="padding:10px 20px;background:${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:10px;color:${isDark ? '#b0b0b0' : '#5d6d7e'};font-size:14px;cursor:pointer;">Cancel</button>
        <button class="block-modal-submit" style="padding:10px 20px;background:linear-gradient(45deg,#e74c3c,#c0392b);border:none;border-radius:10px;color:white;font-size:14px;font-weight:600;cursor:pointer;">Block Selected</button>
      </div>`;

    function renderChips() {
      const list = modal.querySelector('.block-participants-list');
      list.innerHTML = selectedForBlock.map((p, i) => `
        <span class="block-email-chip" data-idx="${i}" style="display:inline-flex;align-items:center;gap:6px;padding:6px 12px;background:${isDark ? 'rgba(231,76,60,0.15)' : 'rgba(231,76,60,0.1)'};color:${isDark ? '#ff8a80' : '#c0392b'};border-radius:20px;font-size:12px;font-weight:500;border:1px solid ${isDark ? 'rgba(231,76,60,0.3)' : 'rgba(231,76,60,0.25)'};transition:all 0.2s;">
          <span style="font-weight:600;">${escapeHtml(p.name)}</span>
          <span style="opacity:0.7;">${escapeHtml(p.email)}</span>
          <span class="block-chip-remove" data-idx="${i}" style="cursor:pointer;font-size:15px;opacity:0.7;line-height:1;">×</span>
        </span>`).join('');
      if (!selectedForBlock.length) {
        list.innerHTML = `<div style="padding:12px;color:${isDark ? '#666' : '#95a5a6'};font-size:13px;text-align:center;width:100%;">No participants selected</div>`;
      }
      list.querySelectorAll('.block-chip-remove').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          selectedForBlock.splice(parseInt(btn.dataset.idx), 1);
          renderChips();
        });
      });
    }
    renderChips();

    // Close handlers
    const closeModal = () => { overlay.style.animation = 'fadeOut 0.2s ease-out'; setTimeout(() => overlay.remove(), 180); };
    modal.querySelector('.block-modal-close').addEventListener('click', closeModal);
    modal.querySelector('.block-modal-cancel').addEventListener('click', closeModal);
    overlay.addEventListener('click', (e) => { if (e.target === overlay) closeModal(); });

    // Submit handler
    modal.querySelector('.block-modal-submit').addEventListener('click', async () => {
      if (!selectedForBlock.length) { closeModal(); return; }
      const submitBtn = modal.querySelector('.block-modal-submit');
      submitBtn.disabled = true;
      submitBtn.textContent = 'Blocking...';
      try {
        await fetch(WEBHOOK_URL, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(createAuthenticatedPayload({
            action: 'add_blocked_participants',
            blocked_participants: selectedForBlock.map(p => p.email),
            todo_id: todoId,
            timestamp: new Date().toISOString(),
            source: 'oracle-chrome-extension-newtab'
          }))
        });
        closeModal();
        showToastNotification(`Blocked ${selectedForBlock.length} participant${selectedForBlock.length > 1 ? 's' : ''}`);
      } catch (err) {
        console.error('Block participants error:', err);
        submitBtn.disabled = false;
        submitBtn.textContent = 'Block Selected';
        showToastNotification('Failed to block participants');
      }
    });

  } catch (err) {
    console.error('Error fetching participants for block:', err);
    overlay.remove();
    showToastNotification('Failed to fetch participants');
  }
}

// --- Blocked Email Participants in Profile ---
function displayBlockedEmailTags(emails) {
  const container = document.getElementById('profileBlockedEmailTags');
  if (!container) return;
  const emailsArray = Array.isArray(emails) ? emails : (emails ? [emails] : []);

  container.innerHTML = emailsArray.map(email => `
    <div class="profile-tag" style="background: linear-gradient(45deg, #e74c3c, #c0392b); border-color: transparent;">
      ${escapeHtml(email)}
      <span class="profile-tag-remove" data-email="${escapeHtml(email)}">×</span>
    </div>
  `).join('');

  container.querySelectorAll('.profile-tag-remove').forEach(btn => {
    btn.addEventListener('click', () => {
      removeBlockedEmail(btn.getAttribute('data-email'));
    });
  });
}

function removeBlockedEmail(email) {
  if (!userProfileData.blocked_email_participants) userProfileData.blocked_email_participants = [];
  const emails = Array.isArray(userProfileData.blocked_email_participants)
    ? userProfileData.blocked_email_participants
    : [userProfileData.blocked_email_participants];
  userProfileData.blocked_email_participants = emails.filter(e => e !== email);
  displayBlockedEmailTags(userProfileData.blocked_email_participants);
  updateSaveButtonVisibility();
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

  // Monitor row button
  const addMonitorBtn = document.getElementById('addMonitorRowBtn');
  if (addMonitorBtn) {
    addMonitorBtn.addEventListener('click', addMonitorRow);
  }

  // Channel input - add on Enter
  if (channelInput) {
    channelInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const channelId = channelInput.value.trim();
        if (channelId) {
          if (!userProfileData.muted_channels) userProfileData.muted_channels = [];
          const channels = Array.isArray(userProfileData.muted_channels) ? userProfileData.muted_channels : [];
          const exists = channels.some(ch => (typeof ch === 'object' ? ch.id : ch) === channelId);
          if (!exists) {
            const name = userProfileData.channel_map?.[channelId] || channelId;
            channels.push({ id: channelId, name });
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
        const channelId = allowedChannelInput.value.trim();
        if (channelId) {
          if (!userProfileData.allowed_public_channels) userProfileData.allowed_public_channels = [];
          const channels = Array.isArray(userProfileData.allowed_public_channels) ? userProfileData.allowed_public_channels : [];
          const exists = channels.some(ch => (typeof ch === 'object' ? ch.id : ch) === channelId);
          if (!exists) {
            const name = userProfileData.channel_map?.[channelId] || channelId;
            channels.push({ id: channelId, name });
            userProfileData.allowed_public_channels = channels;
            displayAllowedChannelTags(channels);
            updateSaveButtonVisibility();
          }
          allowedChannelInput.value = '';
        }
      }
    });
  }

  // Monitor channel input - add on Enter
  const monitorChannelInput = document.getElementById('profileMonitorChannelInput');
  if (monitorChannelInput) {
    monitorChannelInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const channelId = monitorChannelInput.value.trim();
        if (channelId) {
          let channels = userProfileData.monitor_channels || [];
          if (typeof channels === 'string') { try { channels = JSON.parse(channels); } catch { channels = []; } }
          if (!Array.isArray(channels)) channels = [];
          const exists = channels.some(ch => (typeof ch === 'object' ? ch.id : ch) === channelId);
          if (!exists) {
            const name = userProfileData.channel_map?.[channelId] || channelId;
            channels.push({ id: channelId, name });
            userProfileData.monitor_channels = channels;
            displayMonitorChannelTags(channels);
            updateSaveButtonVisibility();
          }
          monitorChannelInput.value = '';
        }
      }
    });
  }

  // Blocked email participant input - add on Enter
  const blockedEmailInput = document.getElementById('profileBlockedEmailInput');
  if (blockedEmailInput) {
    blockedEmailInput.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        const email = blockedEmailInput.value.trim();
        if (email && email.includes('@')) {
          if (!userProfileData.blocked_email_participants) {
            userProfileData.blocked_email_participants = [];
          }
          const emails = Array.isArray(userProfileData.blocked_email_participants)
            ? userProfileData.blocked_email_participants
            : [userProfileData.blocked_email_participants];

          if (!emails.includes(email)) {
            emails.push(email);
            userProfileData.blocked_email_participants = emails;
            displayBlockedEmailTags(emails);
            updateSaveButtonVisibility();
          }
          blockedEmailInput.value = '';
        }
      }
    });
  }
}

// setupProfileOverlay is now called from DOMContentLoaded above
