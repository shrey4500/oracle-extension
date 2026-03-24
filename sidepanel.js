// Oracle Side Panel - Single list (Actions only), no tabs
// Transcript slider on task click, secondary links, no star/edit icons
const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';
const AUTH_URL = 'https://n8n-kqq5.onrender.com/webhook/e6bcd2c3-c714-46c7-94b8-8aeb9831429c';
const CHAT_WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/ffe7e366-078a-4320-aa3b-504458d1d9a4';
// V38: Ably removed — using native n8n webhook streaming
const STORAGE_KEY = 'oracle_user_data';
const READ_TASKS_KEY = 'oracle_read_tasks';
const TASK_TIMESTAMPS_KEY = 'oracle_task_timestamps';

const MEETING_LINK_PATTERNS = [
  'zoom.us', 'zoom.com', 'meet.google.com', 'teams.microsoft.com',
  'teams.live.com', 'webex.com', 'gotomeeting.com', 'bluejeans.com',
  'calendar.google.com', 'google.com/calendar'
];
const DRIVE_LINK_PATTERNS = [
  'docs.google.com/document', 'docs.google.com/spreadsheets',
  'docs.google.com/presentation', 'docs.google.com/forms',
  'drive.google.com', 'sheets.google.com', 'slides.google.com'
];

let allTodos = [], allCalendarItems = [];
let userData = null, isAuthenticated = false;
let readTaskIds = new Set(), isInitialLoad = true;
let previousTaskTimestamps = new Map();
let modifiedMeetingDates = new Set(), activeTagFilters = [];
let pendingActionUpdates = [], pendingActionData = null;
let lastAblyRefreshTime = 0;
const ABLY_REFRESH_THROTTLE_MS = 30000;

// ==================== TAG VALIDATION (matching newtab) ====================
// Read state & validation — delegated to shared components
const { loadReadState, saveReadState, markTaskAsRead, isTaskUnread } = window.Oracle;

// ==================== AUTH ====================
async function initAuth() {
  try { const r = await chrome.storage.local.get([STORAGE_KEY]); if (r[STORAGE_KEY]?.userId) { userData = r[STORAGE_KEY]; isAuthenticated = true; if (window.Oracle && window.Oracle.state) { window.Oracle.state.isAuthenticated = true; window.Oracle.state.userData = userData; } return true; } return false; } catch { return false; }
}
async function login(email, password) {
  try {
    const res = await fetch(AUTH_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ action: 'login', email, password, timestamp: new Date().toISOString() }) });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    let data; try { data = await res.json(); } catch { const t = await res.text(); data = !isNaN(t) ? parseInt(t.trim()) : null; }
    let userId = typeof data === 'number' ? data : typeof data === 'string' && !isNaN(data) ? parseInt(data) : Array.isArray(data) && data[0] ? (data[0].id || data[0].user_id || data[0]) : data?.id || data?.user_id || data?.userId;
    if (!userId) throw new Error('No user ID');
    userData = { userId, email, loginTime: new Date().toISOString() };
    await chrome.storage.local.set({ [STORAGE_KEY]: userData }); isAuthenticated = true; if (window.Oracle && window.Oracle.state) { window.Oracle.state.isAuthenticated = true; window.Oracle.state.userData = userData; }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}
async function logout() { await chrome.storage.local.remove([STORAGE_KEY]); userData = null; isAuthenticated = false; }
function createAuthenticatedPayload(p) { if (!isAuthenticated || !userData) throw new Error('Not authenticated'); return { ...p, user_id: userData.userId, authenticated: true }; }

// ==================== UTILITIES ====================
// Utilities — delegated to shared components
const { escapeHtml, sortTodos, formatDate, formatDueBy, isValidTag } = window.Oracle;
const { isMeetingLink, isDriveLink, isSlackLink, getSlackChannelUrl, extractDriveFileId, getCleanDriveFileUrl } = window.OracleIcons;
function getDateKey(ds) { const d = new Date(ds); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; }

// Recently completed task tracking (prevents tasks reappearing via pending updates)
let recentlyCompletedIds = new Set();
function addRecentlyCompleted(taskId) { const s = String(taskId); recentlyCompletedIds.add(s); setTimeout(() => recentlyCompletedIds.delete(s), 120000); }
function isRecentlyCompleted(taskId) { return recentlyCompletedIds.has(String(taskId)); }

// Icon helpers — delegated to OracleIcons shared component
const { getIconForLink, buildSecondaryLinksHtml: _buildSecLinks } = window.OracleIcons;
// Override buildSecondaryLinksHtml to match sidepanel's simpler version
function buildSecondaryLinksHtml(links) { return _buildSecLinks(links); }
function showLoader(c) { let l = c.querySelector('.loading-overlay'); if (!l) { l = document.createElement('div'); l.className = 'loading-overlay'; l.innerHTML = '<div class="spinner"></div>'; c.appendChild(l); } l.style.display = 'flex'; }
function hideLoader(c) { const l = c.querySelector('.loading-overlay'); if (l) l.style.display = 'none'; }

// ==================== GROUPING (matching newtab.js) ====================
// Grouping — delegated to OracleGrouping shared component
const { groupTasksByDriveFile, groupTasksBySlackChannel, groupTasksByTag } = window.OracleGrouping;

// ==================== LOGIN ====================
function showLoginScreen() {
  const c = document.querySelector('.container'); if (c) c.style.display = 'none';
  document.body.insertAdjacentHTML('beforeend', `<div class="login-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);display:flex;align-items:center;justify-content:center;z-index:10000;"><div style="background:rgba(255,255,255,0.95);backdrop-filter:blur(20px);border-radius:20px;padding:30px;box-shadow:0 20px 40px rgba(0,0,0,0.2);max-width:320px;width:90%;"><div style="text-align:center;margin-bottom:24px;"><div style="width:50px;height:50px;margin:0 auto 12px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:12px;display:flex;align-items:center;justify-content:center;font-size:24px;color:white;font-weight:bold;">∞</div><h2 style="margin:0 0 6px;color:#2c3e50;font-size:20px;">Welcome to Oracle</h2><p style="margin:0;color:#7f8c8d;font-size:13px;">Sign in to continue</p></div><form id="oracleLoginForm" style="display:flex;flex-direction:column;gap:16px;"><input type="email" id="loginEmail" required placeholder="Email" style="width:100%;padding:12px;border:2px solid rgba(225,232,237,0.8);border-radius:10px;font-size:14px;box-sizing:border-box;"><input type="password" id="loginPassword" required placeholder="Password" style="width:100%;padding:12px;border:2px solid rgba(225,232,237,0.8);border-radius:10px;font-size:14px;box-sizing:border-box;"><button type="submit" id="loginSubmitBtn" style="width:100%;background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;padding:12px;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;">Sign In</button><div id="loginError" style="display:none;background:rgba(231,76,60,0.1);border:1px solid rgba(231,76,60,0.3);color:#e74c3c;padding:10px;border-radius:6px;font-size:13px;text-align:center;"></div></form></div></div>`);
  document.getElementById('oracleLoginForm').addEventListener('submit', async (e) => {
    e.preventDefault(); const btn = document.getElementById('loginSubmitBtn'), err = document.getElementById('loginError');
    btn.disabled = true; btn.textContent = 'Signing in...'; err.style.display = 'none';
    const result = await login(document.getElementById('loginEmail').value.trim(), document.getElementById('loginPassword').value.trim());
    if (result.success) { document.querySelector('.login-overlay').remove(); document.querySelector('.container').style.display = 'flex'; initializeMainInterface(); }
    else { err.textContent = result.error; err.style.display = 'block'; }
    btn.disabled = false; btn.textContent = 'Sign In';
  });
}

// ==================== TYPE ICON HELPER ====================
function getTypeIconHtml(type) {
  if (!type) return '';
  const isThread = type === 'slack_thread' || type === 'gmail_email_thread';
  const color = '#667eea';
  if (isThread) {
    return `<span class="task-type-icon" title="Thread" style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;opacity:0.6;flex-shrink:0;color:${color};">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <path d="M9 6h10"/><path d="M9 12h10"/><path d="M9 18h10"/>
        <circle cx="4" cy="6" r="1.5" fill="${color}" stroke="none"/>
        <circle cx="4" cy="12" r="1.5" fill="${color}" stroke="none"/>
        <circle cx="4" cy="18" r="1.5" fill="${color}" stroke="none"/>
      </svg>
    </span>`;
  } else {
    return `<span class="task-type-icon" title="Message" style="display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;opacity:0.6;flex-shrink:0;color:${color};">
      <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="${color}" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
        <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
      </svg>
    </span>`;
  }
}

// ==================== RENDER TODO ITEM (no star, no edit, with secondary links) ====================
function renderTodoItemHtml(todo) {
  const isUnread = isTaskUnread(todo.id) && !isInitialLoad;
  const dueHtml = todo.due_by ? `<span class="todo-due ${new Date(todo.due_by) < new Date() ? 'overdue' : ''}">${formatDueBy(todo.due_by)}</span>` : '';
  const messageLink = todo.message_link || '';
  const primaryIcon = getIconForLink(messageLink);
  const sourceHtml = messageLink ? `<a href="${messageLink}" target="_blank" class="todo-source" title="${primaryIcon.title}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;background:rgba(102,126,234,0.08);transition:all 0.2s;">${primaryIcon.icon}</a>` : '';
  const secondaryHtml = buildSecondaryLinksHtml(todo.secondary_links);
  const todoTags = (todo.tags || []).filter(isValidTag);
  const tagsHtml = todoTags.length > 0 ? `<div class="todo-tags">${todoTags.map(t => `<span class="todo-tag" data-tag="${escapeHtml(t)}">${escapeHtml(t)}</span>`).join('')}</div>` : '';
  const text = todo.task_name || '';
  const maxLen = 120;
  const isLong = text.length > maxLen;
  const displayText = isLong ? escapeHtml(text.substring(0, maxLen)) + '...' : escapeHtml(text);
  const fullTextAttr = isLong ? ` data-full-text="${escapeHtml(text)}"` : '';
  const viewMoreHtml = isLong ? `<span class="view-more-inline" data-todo-id="${todo.id}">View more</span>` : '';
  const participantHtml = todo.participant_text ? `<div class="todo-participant">${escapeHtml(todo.participant_text)}</div>` : '';
  const typeIconHtml = getTypeIconHtml(todo.type || null);

  return `<div class="todo-item${isUnread ? ' unread' : ''}" data-todo-id="${todo.id}" data-message-link="${escapeHtml(messageLink)}">
    <div class="todo-left-actions">
      <div class="todo-checkbox ${todo.status === 1 ? 'checked' : ''}" data-todo-id="${todo.id}">${todo.status === 1 ? '✓' : ''}</div>
    </div>
    <div class="todo-content">
      ${todo.task_title ? `<div class="todo-title">${escapeHtml(todo.task_title)}</div>` : ''}
      <div class="todo-text${isLong ? ' truncated' : ''}"><span class="todo-text-content"${fullTextAttr}>${displayText}</span>${viewMoreHtml}</div>
      ${participantHtml}
      <div class="todo-meta">
        <span class="todo-date">${formatDate(todo.updated_at || todo.created_at)}</span>
        ${dueHtml} ${sourceHtml} ${secondaryHtml}
        ${typeIconHtml ? `<span style="flex:1;min-width:4px;"></span>${typeIconHtml}` : ''}
      </div>
      ${tagsHtml}
    </div>
  </div>`;
}

// ==================== DRIVE GROUP (with doc icon link in header) ====================
function buildTaskGroupHtml(group) {
  const hasUnread = group.tasks.some(t => isTaskUnread(t.id) && !isInitialLoad);
  const docIcon = getIconForLink(group.fileUrl);
  const tasksHtml = group.tasks.map(t => {
    const isUnread = isTaskUnread(t.id) && !isInitialLoad;
    const taskLink = t.message_link || '';
    const pIcon = getIconForLink(taskLink);
    const srcHtml = taskLink ? `<a href="${taskLink}" target="_blank" class="todo-source" title="${pIcon.title}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;background:rgba(102,126,234,0.08);">${pIcon.icon}</a>` : '';
    const secHtml = buildSecondaryLinksHtml(t.secondary_links);
    const todoTags = (t.tags || []).filter(isValidTag);
    const tagsHtml = todoTags.length > 0 ? `<div class="todo-tags" style="margin-top:4px;">${todoTags.map(tg => `<span class="todo-tag" data-tag="${escapeHtml(tg)}">${escapeHtml(tg)}</span>`).join('')}</div>` : '';
    const text = t.task_name || '', maxLen = 120, isLong = text.length > maxLen;
    const displayText = isLong ? escapeHtml(text.substring(0, maxLen)) + '...' : escapeHtml(text);
    const fullAttr = isLong ? ` data-full-text="${escapeHtml(text)}"` : '';
    const vmHtml = isLong ? `<span class="view-more-inline" data-todo-id="${t.id}">View more</span>` : '';
    const tTypeIconHtml = getTypeIconHtml(t.type || null);
    return `<div class="task-group-task-item todo-item${isUnread ? ' unread' : ''}" data-todo-id="${t.id}" data-message-link="${escapeHtml(taskLink)}">
      <div class="todo-left-actions"><div class="todo-checkbox" data-todo-id="${t.id}"></div></div>
      <div class="todo-content" style="flex:1;min-width:0;">
        ${t.task_title ? `<div class="todo-title">${escapeHtml(t.task_title)}</div>` : ''}
        <div class="todo-text${isLong ? ' truncated' : ''}"><span class="todo-text-content"${fullAttr}>${displayText}</span>${vmHtml}</div>
        <div class="todo-meta"><span class="todo-date">${formatDate(t.updated_at || t.created_at)}</span>${srcHtml}${secHtml}${tTypeIconHtml ? `<span style="flex:1;min-width:4px;"></span>${tTypeIconHtml}` : ''}</div>
        ${tagsHtml}
      </div>
    </div>`;
  }).join('');

  return `<div class="task-group${hasUnread ? ' unread' : ''}" data-file-id="${group.fileId}">
    <div class="task-group-header">
      <div class="task-group-checkbox"></div>
      <div class="task-group-icon" style="width:24px;height:24px;display:flex;align-items:center;justify-content:center;">${docIcon.icon}</div>
      <div class="task-group-info">
        <div class="task-group-title">${escapeHtml(group.groupTitle.length > 60 ? group.groupTitle.substring(0, 60) + '...' : group.groupTitle)}</div>
        <div class="task-group-meta"><span class="task-group-count">${group.tasks.length} task${group.tasks.length > 1 ? 's' : ''}</span><span>${formatDate(group.latestUpdate)}</span></div>
      </div>
      <div class="task-group-actions">
        ${group.fileUrl ? `<a href="${group.fileUrl}" target="_blank" class="task-group-open-btn" title="Open document" style="display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;border-radius:6px;background:rgba(102,126,234,0.1);color:#667eea;text-decoration:none;font-size:14px;">↗</a>` : ''}
        <span class="task-group-chevron">▼</span>
      </div>
    </div>
    <div class="task-group-tasks" style="display:none;">${tasksHtml}</div>
  </div>`;
}

// ==================== SLACK GROUP ====================
function buildSlackChannelGroupHtml(group) {
  const hasUnread = group.tasks.some(t => isTaskUnread(t.id) && !isInitialLoad);
  const sortedTasks = [...group.tasks].sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at));
  const slackChannelUrl = getSlackChannelUrl(sortedTasks[0]?.message_link || '');
  const tasksHtml = group.tasks.map(t => {
    const isUnread = isTaskUnread(t.id) && !isInitialLoad;
    const text = t.task_name || '', maxLen = 120, isLong = text.length > maxLen;
    const displayText = isLong ? escapeHtml(text.substring(0, maxLen)) + '...' : escapeHtml(text);
    const fullAttr = isLong ? ` data-full-text="${escapeHtml(text)}"` : '';
    const vmHtml = isLong ? `<span class="view-more-inline" data-todo-id="${t.id}">View more</span>` : '';
    const taskLink = t.message_link || '';
    const pIcon = getIconForLink(taskLink);
    const srcHtml = taskLink ? `<a href="${taskLink}" target="_blank" class="todo-source" title="${pIcon.title}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;background:rgba(102,126,234,0.08);">${pIcon.icon}</a>` : '';
    const secHtml = buildSecondaryLinksHtml(t.secondary_links);
    const todoTags = (t.tags || []).filter(isValidTag);
    const tagsHtml = todoTags.length > 0 ? `<div class="todo-tags" style="margin-top:4px;">${todoTags.map(tg => `<span class="todo-tag" data-tag="${escapeHtml(tg)}">${escapeHtml(tg)}</span>`).join('')}</div>` : '';
    const sTypeIconHtml = getTypeIconHtml(t.type || null);
    return `<div class="slack-channel-task-full todo-item${isUnread ? ' unread' : ''}" data-todo-id="${t.id}" data-message-link="${escapeHtml(taskLink)}">
      <div class="todo-left-actions"><div class="todo-checkbox" data-todo-id="${t.id}"></div></div>
      <div class="todo-content" style="flex:1;min-width:0;">
        ${t.task_title ? `<div class="todo-title">${escapeHtml(t.task_title)}</div>` : ''}
        <div class="todo-text${isLong ? ' truncated' : ''}"><span class="todo-text-content"${fullAttr}>${displayText}</span>${vmHtml}</div>
        <div class="todo-meta"><span class="todo-date">${formatDate(t.updated_at || t.created_at)}</span>${srcHtml}${secHtml}${sTypeIconHtml ? `<span style="flex:1;min-width:4px;"></span>${sTypeIconHtml}` : ''}</div>
        ${tagsHtml}
      </div>
    </div>`;
  }).join('');

  try { var slackIcon = `<img src="${chrome.runtime.getURL('icon-slack.png')}" width="18" height="18">`; } catch (e) { var slackIcon = '💬'; }
  const slackIconHtml = slackChannelUrl
    ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" title="Open Slack channel" style="display:inline-flex;flex-shrink:0;">${slackIcon}</a>`
    : `<span style="display:inline-flex;flex-shrink:0;">${slackIcon}</span>`;
  return `<div class="slack-channel-group${hasUnread ? ' unread' : ''}" data-channel-name="${escapeHtml(group.channelName)}">
    <div class="task-group-header">
      <div class="task-group-checkbox"></div>
      <div class="task-group-icon" style="width:24px;height:24px;display:flex;align-items:center;justify-content:center;">${slackIconHtml}</div>
      <div class="task-group-info">
        <div class="task-group-title">${slackChannelUrl
      ? `<a href="${escapeHtml(slackChannelUrl)}" target="_blank" style="color:inherit;text-decoration:none;">${escapeHtml(group.channelName)}</a>`
      : escapeHtml(group.channelName)}</div>
        <div class="task-group-meta"><span class="task-group-count">${group.tasks.length} task${group.tasks.length > 1 ? 's' : ''}</span><span>${formatDate(group.latestUpdate)}</span></div>
      </div>
      <div class="task-group-actions"><span class="task-group-chevron">▼</span></div>
    </div>
    <div class="task-group-tasks" style="display:none;">${tasksHtml}</div>
  </div>`;
}

function buildTagGroupHtml(group) {
  if (!group.tagName || group.tasks.length === 0) return '';
  const hasUnread = group.tasks.some(t => isTaskUnread(t.id) && !isInitialLoad);
  const sortedTasks = [...group.tasks].sort((a, b) => new Date(b.updated_at || b.created_at) - new Date(a.updated_at || a.created_at));
  const taskIds = sortedTasks.map(t => t.id).join(',');
  const tasksHtml = sortedTasks.map(t => {
    const isUnread = isTaskUnread(t.id) && !isInitialLoad;
    const text = t.task_name || '', maxLen = 120, isLong = text.length > maxLen;
    const displayText = isLong ? escapeHtml(text.substring(0, maxLen)) + '...' : escapeHtml(text);
    const fullAttr = isLong ? ` data-full-text="${escapeHtml(text)}"` : '';
    const vmHtml = isLong ? `<span class="view-more-inline" data-todo-id="${t.id}">View more</span>` : '';
    const taskLink = t.message_link || '';
    const pIcon = getIconForLink(taskLink);
    const srcHtml = taskLink ? `<a href="${taskLink}" target="_blank" class="todo-source" title="${pIcon.title}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;background:rgba(102,126,234,0.08);">${pIcon.icon}</a>` : '';
    const secHtml = buildSecondaryLinksHtml(t.secondary_links);
    const todoTags = (t.tags || []).filter(isValidTag);
    const tagsHtml = todoTags.length > 0 ? `<div class="todo-tags" style="margin-top:4px;">${todoTags.map(tg => `<span class="todo-tag" data-tag="${escapeHtml(tg)}">${escapeHtml(tg)}</span>`).join('')}</div>` : '';
    const tgTypeIconHtml = getTypeIconHtml(t.type || null);
    return `<div class="tag-group-task-item todo-item${isUnread ? ' unread' : ''}" data-todo-id="${t.id}" data-message-link="${escapeHtml(taskLink)}">
      <div class="todo-left-actions"><div class="todo-checkbox" data-todo-id="${t.id}"></div></div>
      <div class="todo-content" style="flex:1;min-width:0;">
        ${t.task_title ? `<div class="todo-title">${escapeHtml(t.task_title)}</div>` : ''}
        <div class="todo-text${isLong ? ' truncated' : ''}"><span class="todo-text-content"${fullAttr}>${displayText}</span>${vmHtml}</div>
        <div class="todo-meta"><span class="todo-date">${formatDate(t.updated_at || t.created_at)}</span>${srcHtml}${secHtml}${tgTypeIconHtml ? `<span style="flex:1;min-width:4px;"></span>${tgTypeIconHtml}` : ''}</div>
        ${tagsHtml}
      </div>
    </div>`;
  }).join('');
  return `<div class="tag-group task-group${hasUnread ? ' unread' : ''}" data-tag-name="${escapeHtml(group.tagName)}" data-task-ids="${taskIds}">
    <div class="task-group-header">
      <div class="task-group-checkbox" data-task-ids="${taskIds}" title="Mark all as done"></div>
      <div class="task-group-icon" style="background:linear-gradient(45deg,#667eea,#764ba2);border-radius:6px;width:24px;height:24px;display:flex;align-items:center;justify-content:center;"><span style="font-size:12px;">🏷️</span></div>
      <div class="task-group-info">
        <div class="task-group-title">${escapeHtml(group.tagName.length > 50 ? group.tagName.substring(0, 50) + '...' : group.tagName)}</div>
        <div class="task-group-meta"><span class="task-group-count">${group.tasks.length} item${group.tasks.length > 1 ? 's' : ''}</span><span>${formatDate(group.latestUpdate)}</span></div>
      </div>
      <div class="task-group-actions"><span class="task-group-chevron">▼</span></div>
    </div>
    <div class="task-group-tasks" style="display:none;">${tasksHtml}</div>
  </div>`;
}

// ==================== MEETINGS ACCORDION ====================
// Format display name: "john.doe" or "john_doe" → "John Doe"
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

// Extract organiser name from participant_text field
function extractOrganiser(participantText) {
  if (!participantText) return '';
  const names = participantText.split(',').map(n => n.trim()).filter(Boolean);
  return names.length > 0 ? formatDisplayName(names[0]) : '';
}

// Show meeting detail slider (sidepanel version - overlay inside content area)
async function showMeetingDetailSlider(meetingId) {
  const meeting = allCalendarItems.find(t => String(t.id) === String(meetingId));
  if (!meeting) return;
  markTaskAsRead(meetingId);

  document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());
  const isDark = document.body.classList.contains('dark-mode');
  const area = document.getElementById('mainContentArea');

  const overlay = document.createElement('div');
  overlay.className = 'transcript-slider-overlay';
  overlay.style.cssText = `position:absolute;top:0;left:0;right:0;bottom:0;background:${isDark ? '#1a1f2e' : '#fff'};z-index:1000;display:flex;flex-direction:column;animation:fadeIn 0.2s ease-out;`;

  const header = document.createElement('div');
  header.style.cssText = `padding:14px 16px;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};display:flex;align-items:center;gap:10px;flex-shrink:0;`;
  header.innerHTML = `
    <div style="flex:1;min-width:0;">
      <div style="font-weight:600;font-size:14px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">📅 Meeting Details</div>
      <div class="meeting-detail-subtitle" style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">Loading...</div>
    </div>
    <button class="transcript-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:32px;height:32px;border-radius:8px;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center;">×</button>`;

  const titleSection = document.createElement('div');
  titleSection.style.cssText = `padding:10px 16px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};flex-shrink:0;`;
  titleSection.innerHTML = `<div style="font-weight:600;font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(meeting.task_title || meeting.task_name || 'Meeting')}</div>`;

  const contentContainer = document.createElement('div');
  contentContainer.style.cssText = 'flex:1;overflow-y:auto;padding:16px;display:flex;flex-direction:column;gap:16px;';
  contentContainer.innerHTML = `
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;">
      <div class="spinner" style="width:32px;height:32px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div>
      <div style="font-size:13px;color:#7f8c8d;">Loading meeting details...</div>
    </div>`;

  overlay.appendChild(header);
  overlay.appendChild(titleSection);
  overlay.appendChild(contentContainer);
  area.appendChild(overlay);

  const closeOverlay = () => { overlay.style.animation = 'fadeOut 0.2s ease-out'; setTimeout(() => overlay.remove(), 200); };
  overlay.querySelector('.transcript-close-btn').addEventListener('click', closeOverlay);
  const escHandlerMeeting = (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); document.removeEventListener('keydown', escHandlerMeeting, true); closeOverlay(); } };
  document.addEventListener('keydown', escHandlerMeeting, true);

  try {
    const res = await fetch(WEBHOOK_URL, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({
        action: 'fetch_task_details',
        todo_id: meetingId,
        message_link: meeting.message_link,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-sidepanel'
      }))
    });
    if (!res.ok) throw new Error('Failed to fetch');
    const responseText = await res.text();
    let data = {};
    if (responseText && responseText.trim()) { try { data = JSON.parse(responseText); } catch (e) { data = {}; } }
    const responseData = Array.isArray(data) ? data[0] : data;
    const md = responseData?.meeting_details || responseData || {};

    const subtitleEl = overlay.querySelector('.meeting-detail-subtitle');
    if (subtitleEl) subtitleEl.textContent = md.date_time ? 'Scheduled' : 'Details loaded';

    let html = '';
    // Date & Time
    const dateTime = md.date_time || meeting.due_by;
    if (dateTime) {
      const dt = new Date(dateTime);
      const dateStr = dt.toLocaleDateString([], { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
      const timeStr = dt.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      const endTime = md.end_time ? new Date(md.end_time).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
      html += `<div style="display:flex;align-items:flex-start;gap:10px;"><div style="width:32px;height:32px;background:rgba(102,126,234,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">🕐</div><div><div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">${dateStr}</div><div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:2px;">${timeStr}${endTime ? ' – ' + endTime : ''}</div></div></div>`;
    }
    // Conference Link
    const confLink = md.conference_link || meeting.message_link || '';
    if (confLink) {
      let label = 'Join Meeting';
      let confIconHtml = `<div style="width:32px;height:32px;background:rgba(46,204,113,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">🔗</div>`;
      try {
        if (confLink.includes('zoom')) { label = 'Join Zoom Meeting'; const iu = chrome.runtime.getURL('icon-zoom.png'); confIconHtml = `<div style="width:32px;height:32px;background:rgba(46,204,113,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;"><img src="${iu}" style="width:20px;height:20px;object-fit:contain;"></div>`; }
        else if (confLink.includes('meet.google')) { label = 'Join Google Meet'; const iu = chrome.runtime.getURL('icon-google-meet.png'); confIconHtml = `<div style="width:32px;height:32px;background:rgba(46,204,113,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;flex-shrink:0;"><img src="${iu}" style="width:20px;height:20px;object-fit:contain;"></div>`; }
        else if (confLink.includes('teams')) { label = 'Join Teams Meeting'; }
      } catch (e) { }
      html += `<div style="display:flex;align-items:flex-start;gap:10px;">${confIconHtml}<div style="flex:1;min-width:0;"><div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Conference</div><a href="${confLink}" target="_blank" style="font-size:11px;color:#667eea;text-decoration:underline;word-break:break-all;display:block;margin-top:2px;">${label}</a></div></div>`;
    }
    // Room
    const room = md.meeting_room || md.location;
    if (room) {
      html += `<div style="display:flex;align-items:flex-start;gap:10px;"><div style="width:32px;height:32px;background:rgba(241,196,15,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">📍</div><div><div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Location / Room</div><div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:2px;">${escapeHtml(room)}</div></div></div>`;
    }
    // Participants
    const participants = md.participants || [];
    if (participants.length > 0) {
      const rows = participants.map(p => {
        const rawName = p.name || p.email?.split('@')[0] || 'Unknown';
        const name = formatDisplayName(rawName);
        const email = p.email || '';
        const status = p.status || p.response_status || 'unknown';
        const isOrg = p.is_organiser || p.organizer || false;
        let sIcon = '❓', sColor = isDark ? '#888' : '#95a5a6', sLabel = status;
        if (status === 'accepted' || status === 'yes') { sIcon = '✅'; sColor = '#27ae60'; sLabel = 'Accepted'; }
        else if (status === 'declined' || status === 'no') { sIcon = '❌'; sColor = '#e74c3c'; sLabel = 'Declined'; }
        else if (status === 'tentative' || status === 'maybe') { sIcon = '❔'; sColor = '#f39c12'; sLabel = 'Tentative'; }
        else if (status === 'needsAction' || status === 'pending' || status === 'awaiting') { sIcon = '⏳'; sColor = isDark ? '#888' : '#95a5a6'; sLabel = 'Pending'; }
        return `<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(225,232,237,0.4)'};"><div style="width:24px;height:24px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:10px;font-weight:600;flex-shrink:0;">${name.charAt(0).toUpperCase()}</div><div style="flex:1;min-width:0;"><div style="font-weight:600;font-size:11px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(name)}${isOrg ? ' <span style="font-size:9px;color:#667eea;">(Organiser)</span>' : ''}</div>${email ? `<div style="font-size:10px;color:${isDark ? '#666' : '#95a5a6'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(email)}</div>` : ''}</div><div style="display:flex;align-items:center;gap:3px;flex-shrink:0;"><span style="font-size:11px;">${sIcon}</span><span style="font-size:10px;color:${sColor};font-weight:500;">${sLabel}</span></div></div>`;
      }).join('');
      html += `<div style="display:flex;align-items:flex-start;gap:10px;"><div style="width:32px;height:32px;background:rgba(155,89,182,0.1);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;flex-shrink:0;">👥</div><div style="flex:1;min-width:0;"><div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};margin-bottom:6px;">Participants (${participants.length})</div><div>${rows}</div></div></div>`;
    }
    if (!html) html = `<div style="text-align:center;padding:30px 16px;color:${isDark ? '#888' : '#95a5a6'};"><div style="font-size:28px;margin-bottom:10px;">📅</div><div style="font-size:13px;">No additional meeting details available.</div></div>`;
    contentContainer.innerHTML = html;
  } catch (e) {
    console.error('Error fetching meeting details:', e);
    contentContainer.innerHTML = `<div style="text-align:center;padding:30px 16px;color:#e74c3c;"><div style="font-size:28px;margin-bottom:10px;">⚠️</div><div style="font-size:13px;">Failed to load meeting details</div></div>`;
  }
}

function buildMeetingsAccordion(meetings) {
  if (!meetings || !meetings.length) return '';
  const sorted = [...meetings].sort((a, b) => { if (!a.due_by && !b.due_by) return 0; if (!a.due_by) return 1; if (!b.due_by) return -1; return new Date(a.due_by) - new Date(b.due_by); });
  const byDate = {};
  sorted.forEach(m => { const dk = m.due_by ? getDateKey(m.due_by) : 'no-date'; if (!byDate[dk]) byDate[dk] = []; byDate[dk].push(m); });
  const dateKeys = Object.keys(byDate).filter(k => k !== 'no-date').sort();
  if (byDate['no-date']?.length) dateKeys.push('no-date');

  const tabsHtml = dateKeys.map((dk, i) => {
    const d = dk !== 'no-date' ? new Date(dk + 'T00:00:00') : null;
    const dayLabel = d ? ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][d.getDay()] : '—';
    const dateLabel = d ? `${d.getDate()}/${d.getMonth() + 1}` : 'TBD';
    return `<div class="meeting-date-tab${i === 0 ? ' active' : ''}" data-date="${dk}"><div class="meeting-date-day">${dayLabel}</div><div class="meeting-date-label">${dateLabel}</div><div class="meeting-date-count">${byDate[dk].length}</div></div>`;
  }).join('');

  const contentHtml = dateKeys.map((dk, i) => {
    const items = byDate[dk].map(m => {
      const time = m.due_by ? new Date(m.due_by).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : '';
      const isOverdue = m.due_by && new Date(m.due_by) < new Date();
      const isUnread = isTaskUnread(m.id) && !isInitialLoad;
      const joinUrl = m.message_link || '';
      const secondaryLinks = m.secondary_links || [];
      const pIcon = getIconForLink(joinUrl);
      let secondaryIconsHtml = '';
      secondaryLinks.forEach(link => {
        if (!link) return;
        let secIcon = null;
        if (link.includes('meet.google.com') || link.includes('zoom.us') || link.includes('zoom.com') || link.includes('teams.microsoft.com')) secIcon = getIconForLink(link);
        else if ((link.includes('calendar.google.com') || link.includes('google.com/calendar')) && !joinUrl.includes('calendar.google.com') && !joinUrl.includes('google.com/calendar')) secIcon = getIconForLink(link);
        if (secIcon) secondaryIconsHtml += `<a href="${link}" target="_blank" class="meeting-secondary-btn" title="${secIcon.title}" style="display:flex;align-items:center;justify-content:center;width:28px;height:28px;background:rgba(102,126,234,0.1);border:1px solid rgba(102,126,234,0.2);border-radius:6px;">${secIcon.icon}</a>`;
      });
      // Extract organiser from participant_text
      const organiserName = extractOrganiser(m.participant_text);
      const timeDisplay = time ? (organiserName ? `${time} · ${formatDueBy(m.due_by)} | ${organiserName}` : `${time} · ${formatDueBy(m.due_by)}`) : (m.due_by ? formatDueBy(m.due_by) : '');
      return `<div class="meeting-item${isUnread ? ' unread' : ''}" data-todo-id="${m.id}" style="cursor:pointer;"><div class="meeting-checkbox" data-todo-id="${m.id}"></div><div class="meeting-info"><div class="meeting-title">${escapeHtml(m.task_title || m.task_name || 'Meeting')}</div><div class="meeting-time${isOverdue ? ' overdue' : ''}">${escapeHtml(timeDisplay)}</div></div><div class="meeting-actions" style="display:flex;align-items:center;gap:6px;">${secondaryIconsHtml}<a href="${joinUrl}" target="_blank" class="meeting-join-btn" title="${pIcon.title}" style="display:flex;align-items:center;justify-content:center;width:28px;height:28px;background:rgba(102,126,234,0.1);border:1px solid rgba(102,126,234,0.2);border-radius:6px;">${pIcon.icon}</a></div></div>`;
    }).join('');
    return `<div class="meetings-date-content${i === 0 ? ' active' : ''}" data-date="${dk}"><div class="meetings-list">${items}</div></div>`;
  }).join('');

  return `<div class="meetings-accordion"><div class="meetings-accordion-header"><div class="meetings-accordion-title"><span>📅</span><span>Meetings</span><span class="meetings-accordion-count">${meetings.length}</span></div><span class="meetings-accordion-chevron">▼</span></div><div class="meetings-accordion-content"><div class="meetings-date-tabs-container"><div class="meetings-date-tabs">${tabsHtml}</div></div>${contentHtml}</div></div>`;
}

// ==================== DISPLAY ACTIONS ====================
function displayActions(todos) {
  const container = document.getElementById('todosContainer');
  if (!container) return;
  hideLoader(container);

  const meetingTodos = todos.filter(t => isMeetingLink(t.message_link) && t.status === 0);
  const nonMeetingTodos = todos.filter(t => !isMeetingLink(t.message_link));
  allCalendarItems = meetingTodos;

  let filtered = nonMeetingTodos.filter(t => t.starred === 1);
  if (activeTagFilters.length > 0) filtered = filtered.filter(t => activeTagFilters.every(ft => (t.tags || []).includes(ft)));
  const sorted = sortTodos(filtered);
  const meetingsHtml = buildMeetingsAccordion(meetingTodos);

  if (sorted.length === 0 && meetingTodos.length === 0) {
    container.innerHTML = `<div class="empty-state"><h3>No starred tasks</h3><p>Star tasks to see them here</p></div>`;
    return;
  }

  const { driveGroups, nonDriveTasks } = groupTasksByDriveFile(sorted);
  const { slackGroups, nonSlackTasks } = groupTasksBySlackChannel(nonDriveTasks);
  const { tagGroups, untaggedTasks } = groupTasksByTag(nonSlackTasks);

  const items = [];
  Object.values(driveGroups).forEach(g => items.push({ type: 'drive-group', data: g, ts: new Date(g.latestUpdate) }));
  Object.values(slackGroups).forEach(g => items.push({ type: 'slack-group', data: g, ts: new Date(g.latestUpdate) }));
  Object.values(tagGroups).forEach(g => items.push({ type: 'tag-group', data: g, ts: new Date(g.latestUpdate) }));
  untaggedTasks.forEach(t => items.push({ type: 'task', data: t, ts: new Date(t.due_by || t.updated_at || t.created_at) }));
  items.sort((a, b) => {
    const aD = a.type === 'task' && a.data.due_by ? new Date(a.data.due_by).getTime() : null;
    const bD = b.type === 'task' && b.data.due_by ? new Date(b.data.due_by).getTime() : null;
    if (aD && bD) return aD - bD; if (aD && !bD) return -1; if (!aD && bD) return 1;
    return b.ts - a.ts;
  });

  let html = meetingsHtml;
  items.forEach(item => {
    if (item.type === 'drive-group') html += buildTaskGroupHtml(item.data);
    else if (item.type === 'slack-group') html += buildSlackChannelGroupHtml(item.data);
    else if (item.type === 'tag-group') html += buildTagGroupHtml(item.data);
    else html += renderTodoItemHtml(item.data);
  });
  container.innerHTML = `<div class="todos-list">${html}</div>`;
  addAllEventListeners(container);
}

// ==================== EVENT LISTENERS ====================
function addAllEventListeners(container) {
  // Meetings accordion
  container.querySelectorAll('.meetings-accordion-header').forEach(h => {
    h.addEventListener('click', () => { const a = h.closest('.meetings-accordion'); a.classList.toggle('open'); if (a.classList.contains('open')) setupMeetingDateTabs(a); });
  });
  setupMeetingDateTabs(container);
  container.querySelectorAll('.meeting-checkbox').forEach(cb => {
    cb.addEventListener('click', (e) => { e.stopPropagation(); if (cb.dataset.todoId) handleCheckbox(cb.dataset.todoId, cb); });
  });

  // Meeting items - click opens meeting detail slider (except checkbox/links)
  container.querySelectorAll('.meeting-item').forEach(item => {
    item.addEventListener('click', (e) => {
      if (e.target.closest('.meeting-checkbox')) return;
      if (e.target.closest('a')) return;
      if (e.target.closest('.meeting-join-btn')) return;
      if (e.target.closest('.meeting-secondary-btn')) return;
      item.classList.remove('unread');
      const meetingId = item.dataset.todoId;
      if (meetingId) showMeetingDetailSlider(meetingId);
    });
  });

  // Task group headers
  container.querySelectorAll('.task-group-header').forEach(header => {
    header.addEventListener('click', (e) => {
      if (e.target.closest('.task-group-checkbox')) { e.stopPropagation(); const g = header.closest('.task-group,.slack-channel-group,.tag-group');[...g.querySelectorAll('[data-todo-id]')].map(el => el.dataset.todoId).filter(Boolean).forEach(id => handleCheckbox(id, e.target)); return; }
      if (e.target.closest('.task-group-open-btn')) return;
      const g = header.closest('.task-group,.slack-channel-group,.tag-group'), tasks = g.querySelector('.task-group-tasks');
      if (g.classList.contains('expanded')) { tasks.style.display = 'none'; g.classList.remove('expanded'); }
      else { tasks.style.display = 'block'; g.classList.add('expanded'); }
    });
  });

  // Todo items - click opens transcript slider (like newtab)
  container.querySelectorAll('.todo-item').forEach(item => {
    const todoId = item.dataset.todoId;
    item.addEventListener('click', (e) => {
      if (e.target.closest('.todo-checkbox') || e.target.closest('.todo-source') || e.target.closest('.todo-tag') || e.target.closest('.view-more-inline') || e.target.closest('a')) return;
      if (todoId) showTranscriptSlider(todoId);
    });
    const cb = item.querySelector('.todo-checkbox');
    if (cb) cb.addEventListener('click', (e) => { e.stopPropagation(); handleCheckbox(cb.dataset.todoId, cb); });
  });

  // View more/less
  container.querySelectorAll('.view-more-inline:not(.view-less)').forEach(btn => {
    attachViewMoreListener(btn);
  });

  // Tags
  container.querySelectorAll('.todo-tag').forEach(el => {
    el.addEventListener('click', (e) => { e.stopPropagation(); const tag = el.dataset.tag; const i = activeTagFilters.indexOf(tag); if (i > -1) activeTagFilters.splice(i, 1); else activeTagFilters.push(tag); displayActions(allTodos); });
  });
}

function attachViewMoreListener(btn) {
  btn.addEventListener('click', (e) => {
    e.stopPropagation();
    const textContent = btn.parentElement.querySelector('.todo-text-content') || btn.previousElementSibling;
    const todoText = btn.closest('.todo-text');
    if (!textContent || !textContent.dataset.fullText) return;
    const fullText = textContent.dataset.fullText;
    const todoId = btn.dataset.todoId;
    markTaskAsRead(todoId);
    // Replace ENTIRE todo-text innerHTML to avoid duplicate buttons
    todoText.innerHTML = `<span class="todo-text-content" data-full-text="${escapeHtml(fullText)}">${escapeHtml(fullText)}</span><span class="view-more-inline view-less" data-todo-id="${todoId}">View less</span>`;
    todoText.classList.remove('truncated'); todoText.classList.add('expanded');
    const viewLessBtn = todoText.querySelector('.view-less');
    if (viewLessBtn) {
      viewLessBtn.addEventListener('click', (ev) => {
        ev.stopPropagation();
        const maxLen = 120;
        todoText.innerHTML = `<span class="todo-text-content" data-full-text="${escapeHtml(fullText)}">${escapeHtml(fullText.substring(0, maxLen))}...</span><span class="view-more-inline" data-todo-id="${todoId}">View more</span>`;
        todoText.classList.remove('expanded'); todoText.classList.add('truncated');
        attachViewMoreListener(todoText.querySelector('.view-more-inline'));
      });
    }
  });
}

function setupMeetingDateTabs(container) {
  container.querySelectorAll('.meeting-date-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      const dk = tab.dataset.date, acc = tab.closest('.meetings-accordion');
      acc.querySelectorAll('.meeting-date-tab').forEach(t => t.classList.remove('active'));
      acc.querySelectorAll('.meetings-date-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      const content = acc.querySelector(`.meetings-date-content[data-date="${dk}"]`);
      if (content) content.classList.add('active');
    });
  });
}

// ==================== DATA LOADING ====================
async function loadTodos() {
  const container = document.getElementById('todosContainer'); if (!container) return;
  showLoader(container); const wasInitialLoad = isInitialLoad;
  try {
    const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'list_todos', filter: 'all', timestamp: new Date().toISOString() })) });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const text = await res.text(); let data = []; if (text?.trim()) try { data = JSON.parse(text); } catch (e) { data = []; }
    allTodos = Array.isArray(data) ? data : (data?.todos || []);
    lastAblyRefreshTime = Date.now();
    if (wasInitialLoad) { allTodos.forEach(t => { readTaskIds.add(String(t.id)); if (t.updated_at) previousTaskTimestamps.set(String(t.id), t.updated_at); }); isInitialLoad = false; window.Oracle.state.isInitialLoad = false; saveReadState(); }
    else { allTodos.forEach(t => { const s = String(t.id); if (t.updated_at) { const prev = previousTaskTimestamps.get(s); if (prev && prev !== t.updated_at) readTaskIds.delete(s); previousTaskTimestamps.set(s, t.updated_at); } }); saveReadState(); }
    hideLoader(container); displayActions(allTodos); updateBadge();
  } catch (e) { console.error('Error:', e); hideLoader(container); container.innerHTML = `<div class="empty-state"><h3>Error: ${e.message}</h3></div>`; }
}
function updateBadge() {
  try {
    const actionUnread = allTodos.filter(t => t.starred === 1 && t.status === 0 && isTaskUnread(t.id) && !isInitialLoad).length;
    const fyiUnread = allTodos.filter(t => t.starred !== 1 && t.status === 0 && isTaskUnread(t.id) && !isInitialLoad).length;
    chrome.runtime.sendMessage({ type: 'updateBadge', actionUnread, fyiUnread, total: actionUnread + fyiUnread });
  } catch (e) { }
}

// ==================== TASK ACTIONS ====================
async function handleCheckbox(id, el) {
  const item = el.closest('.todo-item,.task-group-task-item,.meeting-item,.slack-channel-task-full');
  if (item) {
    item.classList.add('completing');
    item.style.pointerEvents = 'none';
  }
  addRecentlyCompleted(id);
  try {
    await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'toggle_todo', todo_id: id, status: 1, timestamp: new Date().toISOString(), source: 'oracle-sidepanel' })) });
    // Wait for slide-out animation to finish, then silently reload (no loader)
    setTimeout(() => silentReloadTodos(), 500);
  }
  catch (e) { console.error('Error:', e); if (item) { item.classList.remove('completing'); item.style.pointerEvents = ''; } }
}

async function silentReloadTodos() {
  try {
    const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'list_todos', filter: 'all', timestamp: new Date().toISOString() })) });
    if (!res.ok) return;
    const text = await res.text(); let data = []; if (text?.trim()) try { data = JSON.parse(text); } catch (e) { data = []; }
    allTodos = Array.isArray(data) ? data : (data?.todos || []);
    allTodos.forEach(t => { const s = String(t.id); if (t.updated_at) { const prev = previousTaskTimestamps.get(s); if (prev && prev !== t.updated_at) readTaskIds.delete(s); previousTaskTimestamps.set(s, t.updated_at); } });
    saveReadState(); displayActions(allTodos); updateBadge();
  } catch (e) { console.error('Silent reload error:', e); }
}

// ==================== MESSAGE FORMATTING (matching newtab exactly) ====================
// NOTE: formatMessageContent, sanitizeHtml, isComplexEmailHtml, renderEmailInIframe defined below

// ==================== TRANSCRIPT SLIDER (like newtab's showTranscriptSlider) ====================
function showTranscriptSlider(todoId) {
  const todo = [...allTodos, ...allCalendarItems].find(t => String(t.id) === String(todoId));
  if (!todo) return;
  markTaskAsRead(todoId);

  // Remove existing sliders
  document.querySelectorAll('.transcript-slider-overlay').forEach(s => s.remove());

  const isDark = document.body.classList.contains('dark-mode');
  const area = document.getElementById('mainContentArea');

  // Create overlay inside sidepanel content area (not fixed fullscreen)
  const overlay = document.createElement('div');
  overlay.className = 'transcript-slider-overlay';
  overlay.style.cssText = `position:absolute;top:0;left:0;right:0;bottom:0;background:${isDark ? '#1a1f2e' : '#fff'};z-index:1000;display:flex;flex-direction:column;animation:fadeIn 0.2s ease-out;`;

  // Header
  const header = document.createElement('div');
  header.style.cssText = `padding:14px 16px;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};display:flex;align-items:center;gap:10px;flex-shrink:0;`;
  const hdrSourceIcon = getIconForLink(todo.message_link || '');
  header.innerHTML = `
    <div style="flex:1;min-width:0;">
      <div style="font-weight:600;font-size:14px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">💬 Conversation</div>
      <div class="transcript-message-count" style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">Loading...</div>
    </div>
    ${todo.message_link ? `<a href="${todo.message_link}" target="_blank" style="display:inline-flex;align-items:center;justify-content:center;width:32px;height:32px;border-radius:8px;background:rgba(102,126,234,0.1);text-decoration:none;" title="${hdrSourceIcon.title}">${hdrSourceIcon.icon}</a>` : ''}
    <button class="transcript-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:32px;height:32px;border-radius:8px;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center;">×</button>`;
  // Title section
  let titleSection = '';
  if (todo.task_title) {
    titleSection = `<div style="padding:10px 16px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};flex-shrink:0;">
      <div style="font-weight:600;font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(todo.task_title)}</div>
      ${todo.participant_text ? `<div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:3px;">${escapeHtml(todo.participant_text)}</div>` : ''}
    </div>`;
  }

  // Messages container with loading
  const messagesHtml = `<div class="transcript-messages" style="flex:1;overflow-y:auto;overflow-x:hidden;padding:16px;display:flex;flex-direction:column;gap:12px;">
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:12px;">
      <div class="spinner" style="width:32px;height:32px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div>
      <div style="font-size:13px;color:#7f8c8d;">Loading conversation...</div>
    </div>
  </div>`;

  // Reply section
  const isGmail = todo.message_link && todo.message_link.includes('mail.google.com');
  const replyHtml = `<div id="transcriptReplySection" style="padding:16px 20px;border-top:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};background:${isDark ? 'rgba(26,26,46,0.8)' : 'rgba(248,249,250,0.8)'};flex-shrink:0;position:relative;">
    <div class="transcript-attachments" style="display:none;margin-bottom:8px;padding:8px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-radius:8px;border:1px dashed rgba(102,126,234,0.3);">
      <div class="attachments-list" style="display:flex;flex-wrap:wrap;gap:6px;"></div>
    </div>
    <div class="transcript-oracle-overlay" style="display:none;position:absolute;bottom:100%;left:0;right:0;height:120%;min-height:300px;background:${isDark ? '#1a1f2e' : '#f8f9fa'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(200,210,220,0.8)'};border-bottom:none;border-radius:12px 12px 0 0;box-shadow:0 -4px 20px rgba(0,0,0,0.2);z-index:100;flex-direction:column;overflow:hidden;">
      <div class="oracle-overlay-header" style="display:flex;align-items:center;justify-content:space-between;padding:10px 14px;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(200,210,220,0.6)'};background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.12)'};">
        <div style="display:flex;align-items:center;gap:6px;"><div style="width:22px;height:22px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:5px;display:flex;align-items:center;justify-content:center;font-size:12px;color:white;">∞</div><span style="font-weight:600;font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</span></div>
        <button class="oracle-overlay-close" style="background:transparent;border:none;color:${isDark ? '#888' : '#7f8c8d'};cursor:pointer;font-size:16px;padding:4px 6px;border-radius:4px;">×</button>
      </div>
      <div class="oracle-overlay-content" style="flex:1;overflow-y:auto;padding:14px;font-size:13px;line-height:1.6;color:${isDark ? '#e8e8e8' : '#2c3e50'};"></div>
    </div>
    <div class="transcript-recipients-panel" style="display:none;margin-bottom:8px;padding:10px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-radius:8px;border:1px solid ${isDark ? 'rgba(102,126,234,0.3)' : 'rgba(102,126,234,0.15)'};">
      <div class="recipients-section" style="margin-bottom:6px;position:relative;">
        <div style="font-size:10px;font-weight:600;color:${isDark ? '#b0b0b0' : '#5d6d7e'};margin-bottom:3px;">To</div>
        <div style="display:flex;flex-wrap:wrap;gap:4px;align-items:center;padding:5px 8px;background:${isDark ? '#16213e' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:6px;min-height:32px;">
          <div class="recipients-to-chips" style="display:flex;flex-wrap:wrap;gap:4px;"></div>
          <input type="text" class="recipients-to-input" placeholder="Type 3+ chars to search..." style="flex:1;min-width:80px;border:none;background:transparent;outline:none;font-size:11px;color:${isDark ? '#e8e8e8' : '#2c3e50'};padding:2px 0;">
        </div>
        <div class="recipients-to-results" style="display:none;position:absolute;bottom:100%;left:0;right:0;margin-bottom:4px;background:${isDark ? '#1f2940' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'};border-radius:8px;box-shadow:0 -4px 16px rgba(0,0,0,0.15);max-height:150px;overflow-y:auto;z-index:1000;"></div>
      </div>
      <div class="recipients-section" style="margin-bottom:6px;position:relative;">
        <div style="font-size:10px;font-weight:600;color:${isDark ? '#b0b0b0' : '#5d6d7e'};margin-bottom:3px;">CC</div>
        <div style="display:flex;flex-wrap:wrap;gap:4px;align-items:center;padding:5px 8px;background:${isDark ? '#16213e' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:6px;min-height:32px;">
          <div class="recipients-cc-chips" style="display:flex;flex-wrap:wrap;gap:4px;"></div>
          <input type="text" class="recipients-cc-input" placeholder="Type 3+ chars to search..." style="flex:1;min-width:80px;border:none;background:transparent;outline:none;font-size:11px;color:${isDark ? '#e8e8e8' : '#2c3e50'};padding:2px 0;">
        </div>
        <div class="recipients-cc-results" style="display:none;position:absolute;bottom:100%;left:0;right:0;margin-bottom:4px;background:${isDark ? '#1f2940' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'};border-radius:8px;box-shadow:0 -4px 16px rgba(0,0,0,0.15);max-height:150px;overflow-y:auto;z-index:1000;"></div>
      </div>
      <div class="recipients-section" style="position:relative;">
        <div style="font-size:10px;font-weight:600;color:${isDark ? '#b0b0b0' : '#5d6d7e'};margin-bottom:3px;">BCC</div>
        <div style="display:flex;flex-wrap:wrap;gap:4px;align-items:center;padding:5px 8px;background:${isDark ? '#16213e' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:6px;min-height:32px;">
          <div class="recipients-bcc-chips" style="display:flex;flex-wrap:wrap;gap:4px;"></div>
          <input type="text" class="recipients-bcc-input" placeholder="Type 3+ chars to search..." style="flex:1;min-width:80px;border:none;background:transparent;outline:none;font-size:11px;color:${isDark ? '#e8e8e8' : '#2c3e50'};padding:2px 0;">
        </div>
        <div class="recipients-bcc-results" style="display:none;position:absolute;bottom:100%;left:0;right:0;margin-bottom:4px;background:${isDark ? '#1f2940' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.15)' : 'rgba(225,232,237,0.6)'};border-radius:8px;box-shadow:0 -4px 16px rgba(0,0,0,0.15);max-height:150px;overflow-y:auto;z-index:1000;"></div>
      </div>
    </div>
    <div style="display:flex;gap:12px;align-items:flex-end;">
      <div style="display:flex;flex-direction:column;gap:6px;align-items:center;">
        <button class="transcript-mic-btn" title="Voice to text" style="background:rgba(102,126,234,0.1);border:1px solid rgba(102,126,234,0.3);color:#667eea;width:44px;height:38px;border-radius:10px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;transition:all 0.2s;">🎤</button>
        <button class="transcript-recipients-btn" title="Manage Recipients (To/CC/BCC)" style="background:rgba(102,126,234,0.1);border:1px solid rgba(102,126,234,0.3);color:#667eea;width:44px;height:38px;border-radius:10px;cursor:pointer;font-size:16px;display:none;align-items:center;justify-content:center;transition:all 0.2s;">👤</button>
        <input type="file" class="transcript-file-input" multiple style="display:none;">
        <button class="transcript-attach-btn" title="Attach files" style="background:rgba(102,126,234,0.1);border:1px solid rgba(102,126,234,0.3);color:#667eea;width:44px;height:38px;border-radius:10px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;transition:all 0.2s;">📎</button>
      </div>
      <div class="transcript-reply-input" contenteditable="true" placeholder="Type your reply..." style="flex:1;padding:12px 16px;border:2px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:12px;font-size:14px;min-height:80px;max-height:50vh;overflow-y:auto;outline:none;background:${isDark ? '#16213e' : 'white'};color:${isDark ? '#e8e8e8' : '#2c3e50'};line-height:1.5;font-family:inherit;"></div>
      <div style="display:flex;flex-direction:column;gap:6px;align-items:center;">
        <button type="button" class="transcript-oracle-btn" title="Ask Oracle Assistant" style="background:linear-gradient(45deg,#667eea,#764ba2);border:none;color:white;width:34px;height:34px;border-radius:10px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;padding:0;font-size:14px;">∞</button>
        <button class="transcript-send-btn" title="Send (⌘+Enter)" style="background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;width:34px;height:34px;border-radius:10px;font-size:14px;font-weight:600;cursor:pointer;transition:all 0.2s;display:flex;align-items:center;justify-content:center;">↗</button>
        <button class="transcript-done-btn" title="Mark as Done & Close" style="background:rgba(39,174,96,0.12);border:1px solid rgba(39,174,96,0.3);color:#27ae60;width:34px;height:34px;border-radius:10px;font-size:16px;cursor:pointer;transition:all 0.2s;display:flex;align-items:center;justify-content:center;">✓</button>
      </div>
    </div>
    <div class="transcript-sent-timestamp" style="display:none;margin-top:4px;font-size:10px;color:#27ae60;font-weight:500;"></div>
    <div style="margin-top:8px;font-size:11px;color:${isDark ? '#666' : '#95a5a6'};">Press Enter to type, ⌘+Enter to send</div>
  </div>`;

  overlay.appendChild(header);
  overlay.insertAdjacentHTML('beforeend', titleSection);
  overlay.insertAdjacentHTML('beforeend', messagesHtml);
  overlay.insertAdjacentHTML('beforeend', replyHtml);
  area.appendChild(overlay);

  // Close button
  const closeOverlay = () => { overlay.style.animation = 'fadeOut 0.2s ease-out'; document.removeEventListener('keydown', escHandlerTranscript, true); setTimeout(() => overlay.remove(), 200); };
  overlay.querySelector('.transcript-close-btn').addEventListener('click', closeOverlay);
  const escHandlerTranscript = (e) => {
    if (e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation();
      document.removeEventListener('keydown', escHandlerTranscript, true);
      closeOverlay();
    }
  };
  document.addEventListener('keydown', escHandlerTranscript, true);

  // ========== ATTACHMENTS ==========
  const attachBtn = overlay.querySelector('.transcript-attach-btn');
  const fileInput = overlay.querySelector('.transcript-file-input');
  const attachmentsContainer = overlay.querySelector('.transcript-attachments');
  const attachmentsList = overlay.querySelector('.attachments-list');
  let pendingAttachments = [];
  if (attachBtn && fileInput) {
    attachBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
      for (const file of fileInput.files) {
        if (file.size > 10 * 1024 * 1024) { alert(`File ${file.name} is too large (max 10MB)`); continue; }
        try {
          const b64 = await new Promise((res, rej) => { const r = new FileReader(); r.readAsDataURL(file); r.onload = () => res(r.result); r.onerror = rej; });
          pendingAttachments.push({ name: file.name, type: file.type, size: file.size, data: b64 });
          const icon = file.type.startsWith('image/') ? '🖼️' : file.type.includes('pdf') ? '📄' : '📎';
          const sizeStr = file.size < 1024 ? file.size + ' B' : file.size < 1048576 ? (file.size / 1024).toFixed(1) + ' KB' : (file.size / 1048576).toFixed(1) + ' MB';
          const el = document.createElement('div'); el.style.cssText = `display:inline-flex;align-items:center;gap:4px;padding:4px 8px;background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border-radius:6px;font-size:10px;color:${isDark ? '#a0aeff' : '#667eea'};`;
          el.innerHTML = `${icon} <span style="max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(file.name)}</span> <span style="color:${isDark ? '#666' : '#95a5a6'};">(${sizeStr})</span> <button style="background:none;border:none;color:#e74c3c;cursor:pointer;font-size:12px;padding:0 2px;" class="rm-att">×</button>`;
          el.querySelector('.rm-att').addEventListener('click', () => { pendingAttachments = pendingAttachments.filter(a => a.name !== file.name); el.remove(); if (!pendingAttachments.length) attachmentsContainer.style.display = 'none'; });
          attachmentsList.appendChild(el);
        } catch (e) { console.error('File read error:', e); }
      }
      if (pendingAttachments.length) attachmentsContainer.style.display = 'block';
      fileInput.value = '';
    });
  }

  // ========== VOICE-TO-TEXT (MIC) ==========
  const micBtn = overlay.querySelector('.transcript-mic-btn');
  if (micBtn) {
    let recognition = null;
    let isRecording = false;
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      micBtn.style.display = 'none';
    } else {
      micBtn.addEventListener('click', () => {
        if (isRecording && recognition) { recognition.stop(); return; }
        recognition = new SpeechRecognition();
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.lang = 'en-US';
        let finalTranscript = '';
        isRecording = true;
        micBtn.innerHTML = '⏹';
        micBtn.style.background = 'rgba(231,76,60,0.15)';
        micBtn.style.borderColor = 'rgba(231,76,60,0.4)';
        micBtn.style.color = '#e74c3c';
        micBtn.title = 'Stop recording';
        const existingText = replyInput.innerText.trim();
        recognition.onresult = (event) => {
          let interimTranscript = '';
          for (let i = event.resultIndex; i < event.results.length; i++) {
            if (event.results[i].isFinal) { finalTranscript += event.results[i][0].transcript; } else { interimTranscript += event.results[i][0].transcript; }
          }
          const prefix = existingText ? existingText + ' ' : '';
          const interimHtml = interimTranscript ? '<span style="color:#95a5a6;font-style:italic;">' + escapeHtml(interimTranscript) + '</span>' : '';
          replyInput.innerHTML = escapeHtml(prefix + finalTranscript) + interimHtml;
        };
        recognition.onend = () => {
          isRecording = false;
          micBtn.innerHTML = '🎤'; micBtn.style.background = 'rgba(102,126,234,0.1)'; micBtn.style.borderColor = 'rgba(102,126,234,0.3)'; micBtn.style.color = '#667eea'; micBtn.title = 'Voice to text';
          const prefix = existingText ? existingText + ' ' : '';
          if (finalTranscript) { replyInput.innerHTML = escapeHtml(prefix + finalTranscript); }
          recognition = null;
        };
        recognition.onerror = (event) => {
          console.error('Speech recognition error:', event.error);
          isRecording = false;
          micBtn.innerHTML = '🎤'; micBtn.style.background = 'rgba(102,126,234,0.1)'; micBtn.style.borderColor = 'rgba(102,126,234,0.3)'; micBtn.style.color = '#667eea'; micBtn.title = 'Voice to text';
          recognition = null;
        };
        recognition.start();
      });
    }
  }

  // ========== RECIPIENTS PANEL (To/CC/BCC for Gmail) ==========
  const recipientsBtn = overlay.querySelector('.transcript-recipients-btn');
  const recipientsPanel = overlay.querySelector('.transcript-recipients-panel');
  let recipientsPanelOpen = false;
  let recipientTo = [], recipientCc = [], recipientBcc = [];

  function createRecipientChip(person, type) {
    const chip = document.createElement('span');
    const dn = person.name || person.email?.split('@')[0] || 'Unknown';
    chip.style.cssText = `display:inline-flex;align-items:center;gap:3px;background:${isDark ? 'rgba(102,126,234,0.25)' : 'linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1))'};color:${isDark ? '#a0aeff' : '#667eea'};padding:2px 7px;border-radius:12px;font-size:10px;font-weight:500;white-space:nowrap;`;
    chip.dataset.email = person.email || ''; chip.dataset.chipType = type; chip.title = person.email || '';
    chip.innerHTML = `${escapeHtml(dn)} <span class="chip-remove" style="cursor:pointer;font-size:11px;opacity:0.7;margin-left:1px;">×</span>`;
    chip.querySelector('.chip-remove').addEventListener('click', (e) => {
      e.stopPropagation();
      if (type === 'to') recipientTo = recipientTo.filter(r => r.email !== person.email);
      else if (type === 'cc') recipientCc = recipientCc.filter(r => r.email !== person.email);
      else recipientBcc = recipientBcc.filter(r => r.email !== person.email);
      chip.remove();
    });
    return chip;
  }

  async function searchRecipientsFor(query, resultsEl, chipsEl, targetType) {
    if (query.length < 3) { resultsEl.style.display = 'none'; return; }
    resultsEl.style.display = 'block';
    resultsEl.innerHTML = `<div style="padding:8px;text-align:center;font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">Searching...</div>`;
    try {
      const existingToEmails = recipientTo.map(r => r.email).filter(Boolean);
      const existingCcEmails = recipientCc.map(r => r.email).filter(Boolean);
      const existingBccEmails = recipientBcc.map(r => r.email).filter(Boolean);
      const isCC = targetType === 'cc';
      const res = await fetch(WEBHOOK_URL, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({
          action: 'search_user_new_gmail', query, platform: 'gmail',
          existing_to_ids: existingToEmails,
          existing_cc_ids: isCC ? existingCcEmails : existingBccEmails,
          exclude_ids: [...existingToEmails, ...existingCcEmails, ...existingBccEmails],
          timestamp: new Date().toISOString(), source: 'oracle-sidepanel',
          user_id: userData?.userId, authenticated: true
        })
      });
      if (!res.ok) { resultsEl.style.display = 'none'; return; }
      let data = await res.json();
      if (!Array.isArray(data)) data = data.results || data.members || data.users || [];
      // Normalize and filter to people with emails (employee or Direct Message types)
      data = data.filter(r => r.type === 'employee' || r.type === 'Direct Message' || (!r.type && r.user_email_ID));
      data = data.map(r => ({ name: r.Full_Name || r['Full Name'] || r.full_name || r.name || r.user_email_ID || 'Unknown', email: r.user_email_ID || r.email || '', slack_id: r.user_slack_ID || r.slack_id || r.id || '' })).filter(r => r.email);
      if (!data.length) {
        if (query.includes('@')) {
          resultsEl.innerHTML = `<div class="recipient-result" data-email="${escapeHtml(query)}" data-name="${escapeHtml(query.split('@')[0])}" style="padding:7px 10px;cursor:pointer;font-size:11px;color:${isDark ? '#a0aeff' : '#667eea'};">+ Add "${escapeHtml(query)}"</div>`;
        } else { resultsEl.style.display = 'none'; return; }
      } else {
        let manualHtml = '';
        if (query.includes('@')) manualHtml = `<div class="recipient-result" data-email="${escapeHtml(query)}" data-name="${escapeHtml(query.split('@')[0])}" style="padding:7px 10px;cursor:pointer;font-size:11px;color:${isDark ? '#a0aeff' : '#667eea'};border-top:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};">+ Add "${escapeHtml(query)}" manually</div>`;
        resultsEl.innerHTML = data.slice(0, 6).map(u => `<div class="recipient-result" data-name="${escapeHtml(u.name)}" data-email="${escapeHtml(u.email)}" data-slack-id="${escapeHtml(u.slack_id || '')}" style="display:flex;align-items:center;gap:8px;padding:7px 10px;cursor:pointer;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)'};transition:background 0.15s;">
          <div style="width:22px;height:22px;border-radius:50%;background:linear-gradient(135deg,#667eea,#764ba2);display:flex;align-items:center;justify-content:center;color:white;font-size:9px;font-weight:600;flex-shrink:0;">${(u.name || u.email).charAt(0).toUpperCase()}</div>
          <div style="min-width:0;flex:1;"><div style="font-size:11px;font-weight:600;color:${isDark ? '#e0e0e0' : '#2c3e50'};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(u.name || 'Unknown')}</div><div style="font-size:9px;color:${isDark ? '#888' : '#7f8c8d'};overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(u.email)}</div></div>
        </div>`).join('') + manualHtml;
      }
      resultsEl.querySelectorAll('.recipient-result').forEach(r => {
        r.addEventListener('mouseenter', () => r.style.background = isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)');
        r.addEventListener('mouseleave', () => r.style.background = '');
        r.addEventListener('click', () => {
          const p = { name: r.dataset.name, email: r.dataset.email, slack_id: r.dataset.slackId || '' };
          if (targetType === 'to' && !recipientTo.some(x => x.email === p.email)) { recipientTo.push(p); chipsEl.appendChild(createRecipientChip(p, 'to')); }
          else if (targetType === 'cc' && !recipientCc.some(x => x.email === p.email)) { recipientCc.push(p); chipsEl.appendChild(createRecipientChip(p, 'cc')); }
          else if (targetType === 'bcc' && !recipientBcc.some(x => x.email === p.email)) { recipientBcc.push(p); chipsEl.appendChild(createRecipientChip(p, 'bcc')); }
          resultsEl.style.display = 'none'; r.closest('.recipients-section').querySelector('input').value = '';
        });
      });
    } catch (e) { console.error('Recipient search error:', e); resultsEl.style.display = 'none'; }
  }

  // Pre-populate recipients from transcript data (set by fetchTranscript)
  function populateRecipientsFromData() {
    const toParts = overlay._toParticipants || [];
    const ccParts = overlay._ccParticipants || [];
    const allParts = overlay._participants || [];
    const toChipsEl = recipientsPanel.querySelector('.recipients-to-chips');
    const ccChipsEl = recipientsPanel.querySelector('.recipients-cc-chips');
    if (toParts.length > 0 || ccParts.length > 0) {
      toChipsEl.innerHTML = ''; recipientTo = [];
      toParts.forEach(p => { recipientTo.push(p); toChipsEl.appendChild(createRecipientChip(p, 'to')); });
      ccChipsEl.innerHTML = ''; recipientCc = [];
      ccParts.forEach(p => { recipientCc.push(p); ccChipsEl.appendChild(createRecipientChip(p, 'cc')); });
    } else if (allParts.length > 0) {
      toChipsEl.innerHTML = ''; recipientTo = [];
      allParts.forEach(p => { recipientTo.push(p); toChipsEl.appendChild(createRecipientChip(p, 'to')); });
    }
  }

  if (recipientsBtn) {
    let firstOpen = true;
    recipientsBtn.addEventListener('click', () => {
      recipientsPanelOpen = !recipientsPanelOpen;
      recipientsPanel.style.display = recipientsPanelOpen ? 'block' : 'none';
      recipientsBtn.style.background = recipientsPanelOpen ? 'linear-gradient(45deg,#667eea,#764ba2)' : 'rgba(102,126,234,0.1)';
      recipientsBtn.style.color = recipientsPanelOpen ? 'white' : '#667eea';
      recipientsBtn.style.borderColor = recipientsPanelOpen ? 'transparent' : 'rgba(102,126,234,0.3)';
      if (firstOpen && recipientsPanelOpen) { firstOpen = false; populateRecipientsFromData(); }
    });
    // Wire up To/CC/BCC search inputs
    ['to', 'cc', 'bcc'].forEach(type => {
      const inp = recipientsPanel.querySelector(`.recipients-${type}-input`);
      const res = recipientsPanel.querySelector(`.recipients-${type}-results`);
      const chips = recipientsPanel.querySelector(`.recipients-${type}-chips`);
      let t; if (inp) inp.addEventListener('input', () => { clearTimeout(t); t = setTimeout(() => searchRecipientsFor(inp.value.trim(), res, chips, type), 300); });
      if (inp) inp.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault(); const v = inp.value.trim(); if (v && v.includes('@')) {
            const p = { name: v.split('@')[0], email: v };
            if (type === 'to') { recipientTo.push(p); chips.appendChild(createRecipientChip(p, 'to')); }
            else if (type === 'cc') { recipientCc.push(p); chips.appendChild(createRecipientChip(p, 'cc')); }
            else { recipientBcc.push(p); chips.appendChild(createRecipientChip(p, 'bcc')); }
            inp.value = ''; res.style.display = 'none';
          }
        }
      });
    });
  }

  // ========== ORACLE OVERLAY ==========
  const oracleBtn = overlay.querySelector('.transcript-oracle-btn');
  const oracleOverlay = overlay.querySelector('.transcript-oracle-overlay');
  const oracleCloseBtn = overlay.querySelector('.oracle-overlay-close');
  const oracleContent = overlay.querySelector('.oracle-overlay-content');
  if (oracleBtn && oracleOverlay) {
    oracleCloseBtn?.addEventListener('click', () => { oracleOverlay.style.display = 'none'; });
    oracleBtn.addEventListener('click', async () => {
      const messages = overlay.transcriptMessages || [];
      oracleOverlay.style.display = 'flex';
      oracleContent.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:10px;"><div class="spinner" style="width:28px;height:28px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;"></div><div style="color:${isDark ? '#888' : '#7f8c8d'};font-size:12px;">Analyzing thread...</div></div>`;
      const threadConversation = messages.map(msg => ({ role: msg.user || msg.sender || 'unknown', content: msg.text || msg.body || msg.message || '', timestamp: msg.time || msg.timestamp || msg.ts || '' }));
      try {
        const res = await fetch(CHAT_WEBHOOK_URL, {
          method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({
            message: `Analyze this thread and provide insights: ${todo.task_title || todo.task_name || 'Thread'}`,
            sessionId: `transcript_${todoId}_${Date.now()}`, conversation: threadConversation,
            timestamp: new Date().toISOString(), source: 'oracle-transcript',
            context: { task_title: todo.task_title || '', task_name: todo.task_name || '', message_link: todo.message_link || '' }
          })
        });
        if (!res.ok) throw new Error(`HTTP ${res.status}`);
        const result = await res.text(); let rc = result;
        try { const j = JSON.parse(result); rc = j.response || j.message || j.output || result; } catch (e) { }
        let formatted = escapeHtml(rc).replace(/\\n/g, '<br>').replace(/\n/g, '<br>').replace(/^- /gm, '• ');
        oracleContent.innerHTML = `<div style="white-space:pre-wrap;word-wrap:break-word;">${formatted}</div>`;
      } catch (e) {
        oracleContent.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;gap:10px;color:#e74c3c;"><span style="font-size:22px;">⚠️</span><div style="font-size:12px;">Failed to get response from Oracle</div><div style="font-size:10px;color:${isDark ? '#888' : '#95a5a6'};">${escapeHtml(e.message)}</div></div>`;
      }
    });
  }

  // ========== SEND BUTTON ==========
  const sendBtn = overlay.querySelector('.transcript-send-btn');
  const replyInput = overlay.querySelector('.transcript-reply-input');
  const sentTimestamp = overlay.querySelector('.transcript-sent-timestamp');
  const sendReply = async () => {
    const replyHtmlContent = replyInput.innerHTML.trim();
    const msg = replyInput.innerText.trim();
    if (!msg) return;
    // Get Slack-formatted text (converts @mentions to <@SLACK_ID>)
    const getMentionedUsersFromInput = () => {
      const tags = replyInput.querySelectorAll('.mention-tag');
      return Array.from(tags).map(tag => ({ name: tag.textContent.replace('@', ''), email: tag.dataset.email || '', slack_id: tag.dataset.slackId || '' })).filter(u => u.email || u.slack_id);
    };
    const getSlackFormattedText = () => {
      let text = ''; const walk = (node) => { if (node.nodeType === Node.TEXT_NODE) { text += node.textContent; } else if (node.classList && node.classList.contains('mention-tag')) { const sid = node.dataset.slackId; text += sid ? `<@${sid}>` : node.textContent; } else if (node.nodeName === 'BR') { text += '\n'; } else { node.childNodes.forEach(walk); } }; replyInput.childNodes.forEach(walk); return text.trim();
    };
    const mentionedUsersInReply = getMentionedUsersFromInput();
    const slackText = getSlackFormattedText();
    sendBtn.disabled = true; sendBtn.innerHTML = 'Sending...';
    replyInput.innerHTML = '';
    const messagesEl = overlay.querySelector('.transcript-messages');
    // Show sent message on left like other messages (same bubble style)
    messagesEl.insertAdjacentHTML('beforeend', `<div style="display:flex;flex-direction:column;gap:6px;">
      <div style="display:flex;align-items:center;gap:8px;">
        <div style="width:28px;height:28px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:11px;font-weight:600;flex-shrink:0;">You</div>
        <div style="flex:1;min-width:0;"><div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">You</div><div style="font-size:10px;color:${isDark ? '#666' : '#95a5a6'};">Just now</div></div>
      </div>
      <div style="margin-left:36px;padding:10px 14px;background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border-radius:12px;border-top-left-radius:4px;font-size:13px;color:${isDark ? '#e0e0e0' : '#2c3e50'};line-height:1.6;">${escapeHtml(msg)}</div>
    </div>`);
    messagesEl.scrollTop = messagesEl.scrollHeight;
    try {
      const isDrive = isDriveLink(todo.message_link);
      const replyAction = isDrive ? 'reply_drive_comment' : 'reply_to_thread';
      await fetch(WEBHOOK_URL, {
        method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({
          action: replyAction,
          todo_id: todoId,
          message_link: todo.message_link,
          reply_text: isDrive ? msg : slackText,
          reply_text_plain: msg,
          reply_html: replyHtmlContent,
          mentioned_users: mentionedUsersInReply.length > 0 ? mentionedUsersInReply : undefined,
          to_emails: recipientTo.length > 0 ? recipientTo.map(r => r.email).filter(Boolean) : undefined,
          cc_emails: recipientCc.length > 0 ? recipientCc.map(r => r.email).filter(Boolean) : undefined,
          bcc_emails: recipientBcc.length > 0 ? recipientBcc.map(r => r.email).filter(Boolean) : undefined,
          attachments: pendingAttachments.length > 0 ? pendingAttachments : undefined,
          timestamp: new Date().toISOString(),
          source: 'oracle-chrome-extension-sidepanel'
        }))
      });
      sendBtn.style.background = 'linear-gradient(45deg,#27ae60,#2ecc71)'; sendBtn.innerHTML = '✓';
      mentionedUsers = [];
      pendingAttachments = []; if (attachmentsList) attachmentsList.innerHTML = ''; if (attachmentsContainer) attachmentsContainer.style.display = 'none';
      // Show sent timestamp
      const now = new Date(); const timeStr = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
      if (sentTimestamp) { sentTimestamp.textContent = `✓ Sent at ${timeStr}`; sentTimestamp.style.display = 'block'; }
      // Close recipients panel
      if (recipientsPanelOpen && recipientsPanel) { recipientsPanelOpen = false; recipientsPanel.style.display = 'none'; if (recipientsBtn) { recipientsBtn.style.background = 'rgba(102,126,234,0.1)'; recipientsBtn.style.color = '#667eea'; } }
      setTimeout(() => { sendBtn.style.background = 'linear-gradient(45deg,#667eea,#764ba2)'; sendBtn.innerHTML = '↗'; sendBtn.disabled = false; }, 1500);
      // Refresh transcript after 2s
      setTimeout(() => fetchTranscript(todoId, overlay, todo), 2000);
    } catch (e) { console.error('Send error:', e); sendBtn.innerHTML = '↗'; sendBtn.disabled = false; }
  };
  sendBtn.addEventListener('click', sendReply);

  // Mark Done button handler — marks task as done and closes slider
  const doneBtn = overlay.querySelector('.transcript-done-btn');
  if (doneBtn) {
    doneBtn.addEventListener('click', () => {
      // Immediately close the slider
      closeOverlay();

      // Remove the task row from the DOM
      const taskRow = document.querySelector(`.todo-item[data-todo-id="${todo.id}"]`);
      if (taskRow) {
        const checkbox = taskRow.querySelector('.todo-checkbox');
        if (checkbox) { checkbox.classList.add('checked'); checkbox.innerHTML = '✓'; }
        taskRow.classList.add('completing');
        setTimeout(() => taskRow.remove(), 400);
      }

      // Update local arrays
      if (typeof allTodos !== 'undefined') allTodos = allTodos.filter(t => t.id != todo.id);
      if (typeof allFyiItems !== 'undefined') allFyiItems = allFyiItems.filter(t => t.id != todo.id);

      // Send to backend in background
      fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'toggle_todo', todo_id: todo.id, status: 1, timestamp: new Date().toISOString(), source: 'oracle-sidepanel-slider' })) }).catch(err => console.error('Error marking task done:', err));
    });
    doneBtn.addEventListener('mouseenter', () => { doneBtn.style.background = 'rgba(39,174,96,0.22)'; doneBtn.style.transform = 'scale(1.05)'; });
    doneBtn.addEventListener('mouseleave', () => { doneBtn.style.background = 'rgba(39,174,96,0.12)'; doneBtn.style.transform = 'scale(1)'; });
  }

  replyInput.addEventListener('keydown', (e) => {
    if ((e.metaKey || e.ctrlKey) && e.key === 'Enter') { e.preventDefault(); sendReply(); return; }
    if (e.key === 'Enter' && !e.ctrlKey && !e.metaKey && !e.shiftKey) { e.preventDefault(); document.execCommand('insertLineBreak'); }
    if (e.key === 'Escape') {
      e.preventDefault(); e.stopPropagation();
      // If mention dropdown is open, close it only
      if (mentionDropdown && mentionDropdown.style.display !== 'none') { hideMentionDropdown(); return; }
      // Otherwise close the whole slider
      document.removeEventListener('keydown', escHandlerTranscript, true);
      closeOverlay();
    }
  });

  // ===== @mention autocomplete =====
  let mentionDropdown = null, mentionSearchTimeout = null, mentionSelectedIndex = 0, mentionResults = [], currentMentionQuery = '', mentionedUsers = [];
  const replySection = overlay.querySelector('#transcriptReplySection');
  console.log('[Oracle] replySection found:', !!replySection);

  const createMentionDropdown = () => {
    if (mentionDropdown) return mentionDropdown;
    mentionDropdown = document.createElement('div');
    mentionDropdown.className = 'mention-dropdown';
    mentionDropdown.style.cssText = `display:none;position:absolute;bottom:100%;left:0;right:0;max-height:200px;overflow-y:auto;background:${isDark ? '#1f2940' : '#fff'};border:1px solid ${isDark ? 'rgba(255,255,255,0.15)' : 'rgba(200,210,220,0.8)'};border-radius:10px;box-shadow:0 -4px 16px rgba(0,0,0,${isDark ? '0.4' : '0.12'});z-index:100;`;
    if (replySection) { replySection.insertBefore(mentionDropdown, replySection.firstChild); }
    else { overlay.appendChild(mentionDropdown); console.warn('[Oracle] replySection null, appended to overlay'); }
    mentionDropdown.addEventListener('mousedown', (e) => {
      const item = e.target.closest('.mention-item');
      if (item) { e.preventDefault(); e.stopPropagation(); insertMention(item.dataset.name, item.dataset.email, item.dataset.slackId || ''); }
    });
    return mentionDropdown;
  };
  const showMentionLoading = () => { const d = createMentionDropdown(); d.innerHTML = `<div style="padding:10px;text-align:center;font-size:12px;color:${isDark ? '#888' : '#7f8c8d'};">Searching...</div>`; d.style.display = 'block'; };
  const showMentionResults = (users) => {
    const d = createMentionDropdown(); mentionResults = users; mentionSelectedIndex = 0;
    if (!users.length) { d.innerHTML = `<div style="padding:10px;text-align:center;font-size:12px;color:${isDark ? '#888' : '#7f8c8d'};">No users found</div>`; d.style.display = 'block'; return; }
    d.innerHTML = users.map((u, i) => {
      const initials = (u.name || u.email || '?').split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
      return `<div class="mention-item${i === 0 ? ' selected' : ''}" data-index="${i}" data-name="${escapeHtml(u.name || u.email)}" data-email="${escapeHtml(u.email || '')}" data-slack-id="${escapeHtml(u.slack_id || '')}" style="display:flex;align-items:center;gap:10px;padding:8px 12px;cursor:pointer;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.05)'};${i === 0 ? `background:${isDark ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.08)'}` : ''};"><div style="width:28px;height:28px;border-radius:50%;background:linear-gradient(135deg,#667eea,#764ba2);display:flex;align-items:center;justify-content:center;color:white;font-size:10px;font-weight:600;flex-shrink:0;">${initials}</div><div style="min-width:0;"><div style="font-size:12px;font-weight:600;color:${isDark ? '#e0e0e0' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(u.name || 'Unknown')}</div>${u.email ? `<div style="font-size:10px;color:${isDark ? '#888' : '#7f8c8d'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(u.email)}</div>` : ''}</div></div>`;
    }).join('');
    d.style.display = 'block';
  };
  const hideMentionDropdown = () => { if (mentionDropdown) mentionDropdown.style.display = 'none'; mentionResults = []; mentionSelectedIndex = 0; currentMentionQuery = ''; };
  const updateMentionSelection = () => { if (!mentionDropdown) return; mentionDropdown.querySelectorAll('.mention-item').forEach((item, i) => { item.style.background = i === mentionSelectedIndex ? (isDark ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.08)') : ''; item.classList.toggle('selected', i === mentionSelectedIndex); if (i === mentionSelectedIndex) item.scrollIntoView({ block: 'nearest' }); }); };
  const insertMention = (name, email, slackId) => {
    if (email && !mentionedUsers.find(u => u.email === email)) mentionedUsers.push({ name, email, slack_id: slackId });
    const span = document.createElement('span');
    span.className = 'mention-tag'; span.contentEditable = 'false'; span.dataset.email = email || ''; span.dataset.slackId = slackId || '';
    span.textContent = `@${name}`;
    span.style.cssText = `background:linear-gradient(45deg,rgba(102,126,234,0.2),rgba(118,75,162,0.2));color:#667eea;padding:2px 6px;border-radius:4px;font-weight:600;font-size:12px;display:inline-block;margin:0 2px;`;
    const sel = window.getSelection(); let replaced = false;
    if (sel.rangeCount > 0) { const r = sel.getRangeAt(0); let cn = r.endContainer; if (cn.nodeType === Node.TEXT_NODE && !(cn.parentElement && cn.parentElement.classList.contains('mention-tag'))) { const tc = cn.textContent, co = r.endOffset; let atPos = -1; for (let i = co - 1; i >= 0; i--) { if (tc[i] === '@') { atPos = i; break; } if (tc[i] === ' ' || tc[i] === '\n') break; } if (atPos !== -1) { const before = tc.substring(0, atPos), after = tc.substring(co), parent = cn.parentNode, frag = document.createDocumentFragment(); if (before) frag.appendChild(document.createTextNode(before)); frag.appendChild(span); frag.appendChild(document.createTextNode('\u00A0' + after)); parent.replaceChild(frag, cn); replaced = true; } } }
    if (!replaced) { replyInput.appendChild(span); replyInput.appendChild(document.createTextNode('\u00A0')); }
    const nr = document.createRange(); const ns = window.getSelection(); nr.selectNodeContents(replyInput); nr.collapse(false); ns.removeAllRanges(); ns.addRange(nr);
    mentionInsertCooldown = true; setTimeout(() => { mentionInsertCooldown = false; }, 500);
    hideMentionDropdown(); replyInput.focus();
  };
  let mentionAbortController = null;
  const searchMentionUsers = async (query) => {
    if (mentionAbortController) { mentionAbortController.abort(); mentionAbortController = null; }
    currentMentionQuery = query; showMentionLoading();
    console.log('[Oracle] Searching mentions for:', query);
    const controller = new AbortController();
    mentionAbortController = controller;
    const timeoutId = setTimeout(() => controller.abort(), 10000);
    try {
      const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'search_user', query, timestamp: new Date().toISOString() })), signal: controller.signal });
      clearTimeout(timeoutId);
      mentionAbortController = null;
      if (!res.ok) throw new Error('Search failed');
      const rt = await res.text(); let data = [];
      if (rt && rt.trim()) { try { data = JSON.parse(rt); } catch (e) { data = []; } }
      if (typeof data === 'string') { try { data = JSON.parse(data); } catch (e) { data = []; } }
      let users = Array.isArray(data) ? data : (data?.users || data?.results || []);
      users = users.map(u => ({ name: u['Full Name'] || u.name || u.full_name || '', email: u.user_email_ID || u.email || '', slack_id: u.user_slack_ID || u.slack_id || '' })).filter(u => u.name || u.email);
      if (query === currentMentionQuery) showMentionResults(users);
    } catch (e) { clearTimeout(timeoutId); mentionAbortController = null; if (e.name !== 'AbortError') console.error('Mention search error:', e); if (query === currentMentionQuery) hideMentionDropdown(); }
  };
  let mentionInsertCooldown = false;
  const checkForMention = () => {
    if (mentionInsertCooldown) return;
    const sel = window.getSelection(); if (!sel.rangeCount) return;
    const r = sel.getRangeAt(0), cn = r.endContainer;
    if (cn.nodeType !== Node.TEXT_NODE || (cn.parentElement && cn.parentElement.classList.contains('mention-tag'))) { hideMentionDropdown(); return; }
    const tc = cn.textContent, co = r.endOffset;
    let atPos = -1; for (let i = co - 1; i >= 0; i--) { if (tc[i] === '@') { atPos = i; break; } if (tc[i] === ' ' || tc[i] === '\u00A0' || tc[i] === '\n') break; }
    if (atPos !== -1) {
      const query = tc.substring(atPos + 1, co);
      if (query.length >= 3 && !/\s/.test(query)) {
        if (query === currentMentionQuery && mentionDropdown && mentionDropdown.style.display === 'block') return;
        if (mentionSearchTimeout) clearTimeout(mentionSearchTimeout);
        mentionSearchTimeout = setTimeout(() => searchMentionUsers(query), 400);
      }
      else if (query.length < 3) hideMentionDropdown();
    } else hideMentionDropdown();
  };
  replyInput.addEventListener('input', checkForMention);
  replyInput.addEventListener('keydown', (e) => {
    if (!mentionDropdown || mentionDropdown.style.display === 'none' || !mentionResults.length) return;
    if (e.key === 'ArrowDown') { e.preventDefault(); mentionSelectedIndex = Math.min(mentionSelectedIndex + 1, mentionResults.length - 1); updateMentionSelection(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); mentionSelectedIndex = Math.max(mentionSelectedIndex - 1, 0); updateMentionSelection(); }
    else if (e.key === 'Enter' || e.key === 'Tab') { if (mentionResults.length > 0) { e.preventDefault(); const s = mentionResults[mentionSelectedIndex]; insertMention(s.name || s.email, s.email || '', s.slack_id || ''); } }
    else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); hideMentionDropdown(); }
  });

  // Fetch transcript
  fetchTranscript(todoId, overlay, todo);
}

// ==================== MESSAGE FORMATTING (from newtab) ====================
function formatMessageContent(text) {
  if (!text) return '';
  const isDark = document.body.classList.contains('dark-mode');
  const textColor = isDark ? '#e8e8e8' : '#2c3e50';
  if (/<[a-z][\s\S]*>/i.test(text)) return sanitizeHtml(text);
  let f = text;
  f = f.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  f = escapeHtml(f);
  // Emoji map (common ones)
  const em = { ':slightly_smiling_face:': '🙂', ':smile:': '😄', ':grinning:': '😀', ':laughing:': '😆', ':blush:': '😊', ':wink:': '😉', ':thinking_face:': '🤔', ':thinking:': '🤔', ':raised_hands:': '🙌', ':clap:': '👏', ':pray:': '🙏', ':thumbsup:': '👍', ':+1:': '👍', ':-1:': '👎', ':ok_hand:': '👌', ':wave:': '👋', ':fire:': '🔥', ':star:': '⭐', ':sparkles:': '✨', ':heart:': '❤️', ':100:': '💯', ':warning:': '⚠️', ':white_check_mark:': '✅', ':x:': '❌', ':heavy_check_mark:': '✔️', ':eyes:': '👀', ':rocket:': '🚀', ':tada:': '🎉', ':bulb:': '💡', ':memo:': '📝', ':pushpin:': '📌', ':calendar:': '📅', ':email:': '📧', ':phone:': '📞', ':computer:': '💻', ':link:': '🔗', ':lock:': '🔒', ':chart_with_upwards_trend:': '📈', ':bar_chart:': '📊', ':moneybag:': '💰', ':trophy:': '🏆', ':muscle:': '💪', ':brain:': '🧠', ':robot:': '🤖', ':zap:': '⚡', ':boom:': '💥', ':handshake:': '🤝', ':red_circle:': '🔴', ':large_blue_circle:': '🔵', ':green_circle:': '🟢', ':arrow_right:': '➡️', ':arrow_left:': '⬅️' };
  Object.keys(em).forEach(code => { f = f.split(code).join(em[code]); });
  f = f.replace(/:([a-z0-9_+-]+):/g, (m, n) => em[m] || m);
  // Code blocks (triple backticks)
  f = f.replace(/```([^`]+)```/g, '<pre style="background:rgba(45,55,72,0.08);border:1px solid rgba(45,55,72,0.15);border-radius:6px;padding:10px;margin:6px 0;font-family:monospace;font-size:11px;overflow-x:hidden;white-space:pre-wrap;word-break:break-all;max-width:100%;">$1</pre>');
  // Inline code
  f = f.replace(/`([^`\n]+)`/g, `<code style="background:rgba(45,55,72,0.08);border-radius:4px;padding:2px 6px;font-family:monospace;font-size:11px;color:#e53e3e;">$1</code>`);
  // Markdown links [text](url)
  f = f.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, '<a href="$2" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$1</a>');
  // Slack <url|text>
  f = f.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$2</a>');
  // Slack <url>
  f = f.replace(/<(https?:\/\/[^>]+)>/g, '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$1</a>');
  // Plain URLs
  f = f.replace(/(?<!href="|">)(https?:\/\/[^\s<>)\]]+)/g, '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">$1</a>');
  // Email addresses
  f = f.replace(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/g, '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$1</a>');
  // Bold *text*
  f = f.replace(/\*([^*\n]+)\*/g, `<strong style="font-weight:600;color:${textColor};">$1</strong>`);
  // Italic _text_
  f = f.replace(/(?<!\w)_([^_\n]+)_(?!\w)/g, '<em style="font-style:italic;">$1</em>');
  // Strikethrough ~text~
  f = f.replace(/~([^~\n]+)~/g, '<del style="text-decoration:line-through;opacity:0.7;">$1</del>');
  // Numbered list items
  f = f.replace(/^(\d+)\.\s+/gm, '<span style="color:#667eea;font-weight:600;margin-right:4px;">$1.</span> ');
  // Remove [image: ...] placeholders
  f = f.replace(/\[image:[^\]]+\]/g, '');
  // Blockquotes (> lines)
  const lines = f.split('<br>'); let inQ = false, qLines = [], res = [];
  for (let i = 0; i < lines.length; i++) { const l = lines[i]; if (/^(&gt;|>)\s?/.test(l.trim())) { qLines.push(l.trim().replace(/^(&gt;|>)\s?/, '')); inQ = true; } else { if (inQ && qLines.length) { res.push(`<blockquote style="border-left:4px solid #667eea;margin:6px 0;padding:6px 10px;background:rgba(102,126,234,0.08);border-radius:0 6px 6px 0;color:inherit;font-style:normal;">${qLines.join('<br>')}</blockquote>`); qLines = []; inQ = false; } res.push(l); } }
  if (qLines.length) res.push(`<blockquote style="border-left:4px solid #667eea;margin:6px 0;padding:6px 10px;background:rgba(102,126,234,0.08);border-radius:0 6px 6px 0;color:inherit;font-style:normal;">${qLines.join('<br>')}</blockquote>`);
  f = res.join('<br>');
  f = f.replace(/<br>{3,}/g, '<br><br>');
  const preParts = f.split(/(<pre[^>]*>[\s\S]*?<\/pre>)/);
  f = preParts.map(p => p.startsWith('<pre') ? p : p.replace(/\n/g, '<br>')).join('');
  return f;
}

function sanitizeHtml(html) {
  const isDark = document.body.classList.contains('dark-mode');
  const temp = document.createElement('div'); temp.innerHTML = html;
  temp.querySelectorAll('script,style,iframe,object,embed').forEach(el => el.remove());
  temp.querySelectorAll('*').forEach(el => { Array.from(el.attributes).forEach(attr => { if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) el.removeAttribute(attr.name); }); });
  temp.querySelectorAll('table').forEach(t => { t.style.cssText = 'border-collapse:collapse;width:100%;margin:8px 0;font-size:12px;overflow-x:hidden;display:block;max-width:100%;table-layout:fixed;'; });
  temp.querySelectorAll('th').forEach(th => { th.style.cssText = 'background:linear-gradient(45deg,#667eea,#764ba2);color:white;padding:8px 10px;text-align:left;font-weight:600;border:1px solid rgba(102,126,234,0.3);white-space:nowrap;'; });
  temp.querySelectorAll('td').forEach(td => { td.style.cssText = `padding:6px 10px;border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};vertical-align:top;`; });
  temp.querySelectorAll('tr').forEach((tr, i) => { if (i % 2 === 1) tr.style.backgroundColor = 'rgba(102,126,234,0.03)'; });
  temp.querySelectorAll('a').forEach(link => { link.style.cssText = 'color:#667eea;text-decoration:underline;'; link.setAttribute('target', '_blank'); });
  temp.querySelectorAll('p').forEach(p => { p.style.cssText = 'margin:6px 0;line-height:1.5;'; });
  temp.querySelectorAll('b,strong').forEach(b => { b.style.cssText = `font-weight:600;color:${isDark ? '#e8e8e8' : '#2c3e50'};`; });
  return temp.innerHTML;
}

function isComplexEmailHtml(html) {
  if (!html) return false;
  const temp = document.createElement('div');
  temp.innerHTML = html;
  const nestedTables = temp.querySelectorAll('table table');
  const signatureBlocks = temp.querySelectorAll('[id*="Signature"], [id*="signature"], .elementToProof');
  const hrTags = temp.querySelectorAll('hr');
  const hasMultipleInlineStyles = (html.match(/style="/g) || []).length > 10;
  return nestedTables.length > 0 || signatureBlocks.length > 0 || (hrTags.length > 0 && hasMultipleInlineStyles);
}

function renderEmailInIframe(html, container, isDarkMode) {
  const temp = document.createElement('div');
  temp.innerHTML = html;
  temp.querySelectorAll('script, object, embed').forEach(el => el.remove());
  temp.querySelectorAll('*').forEach(el => {
    Array.from(el.attributes).forEach(attr => {
      if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) el.removeAttribute(attr.name);
    });
    const style = el.getAttribute('style') || '';
    if (style) {
      el.setAttribute('style', style.replace(/min-width\s*:\s*\d{3,}px/gi, 'min-width:0'));
    }
  });
  temp.querySelectorAll('a').forEach(link => { link.setAttribute('target', '_blank'); link.setAttribute('rel', 'noopener noreferrer'); });
  const cleanHtml = temp.innerHTML;
  const textColor = isDarkMode ? '#e8e8e8' : '#2c3e50';
  const linkColor = isDarkMode ? '#8b9ff0' : '#667eea';
  const iframeDoc = `<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>
*{box-sizing:border-box}
html{overflow-x:auto;overflow-y:auto}
body{margin:0;padding:6px 10px;background:transparent;color:${textColor};font-family:'Lato',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:13px;line-height:1.5;word-wrap:break-word;overflow-wrap:break-word;overflow-x:auto}
a{color:${linkColor};word-break:break-all}
img{max-width:100%!important;height:auto!important}
table{border-collapse:collapse;overflow:visible}
td,th{word-break:normal;overflow-wrap:break-word}
div,span,p,section,article{overflow-wrap:break-word}
pre{white-space:pre-wrap;word-break:break-all;overflow-x:auto;max-width:100%}
blockquote{margin:8px 0;padding:8px 12px;border-left:3px solid ${linkColor};background:${isDarkMode ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};max-width:100%;overflow:auto}
hr{border:none;border-top:1px solid ${isDarkMode ? 'rgba(255,255,255,0.12)' : 'rgba(0,0,0,0.08)'};margin:10px 0}
table[cellpadding],table[cellspacing]{font-size:12px}
h2{font-size:14px;margin:4px 0}
p{margin:4px 0}
ul,ol{margin:4px 0;padding-left:20px}
li{margin:2px 0}
b,strong{font-weight:600;color:${textColor}}
img[alt="mobilePhone"],img[alt="emailAddress"],img[alt="website"],img[alt="address"]{width:12px!important;height:auto}
div[class*="elementToProof"]{margin-top:0.3em;margin-bottom:0.3em}
</style></head><body>${cleanHtml}</body></html>`;
  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'width:100%;border:none;min-height:50px;background:transparent;display:block;margin:0;padding:0;overflow:hidden;';
  iframe.setAttribute('sandbox', 'allow-same-origin allow-popups allow-popups-to-escape-sandbox');
  iframe.srcdoc = iframeDoc;
  iframe.onload = () => {
    try {
      const resizeIframe = () => {
        if (iframe.contentDocument && iframe.contentDocument.body) {
          iframe.style.height = Math.min(iframe.contentDocument.body.scrollHeight + 4, 1500) + 'px';
        }
      };
      resizeIframe();
      iframe.contentDocument.querySelectorAll('img').forEach(img => { if (!img.complete) img.addEventListener('load', resizeIframe); });
      setTimeout(resizeIframe, 300);
      setTimeout(resizeIframe, 1000);
    } catch (e) { console.warn('Could not auto-resize email iframe:', e); }
  };
  container.style.padding = '0';
  container.style.background = 'transparent';
  container.innerHTML = '';
  container.appendChild(iframe);
}

async function fetchTranscript(todoId, overlay, todo) {
  const messagesEl = overlay.querySelector('.transcript-messages');
  const countEl = overlay.querySelector('.transcript-message-count');
  const isDark = document.body.classList.contains('dark-mode');
  try {
    // Use same API call as newtab: fetch_task_details or fetch_comment_details for Drive
    const fetchAction = isDriveLink(todo.message_link) ? 'fetch_comment_details' : 'fetch_task_details';
    const res = await fetch(WEBHOOK_URL, {
      method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({
        action: fetchAction,
        todo_id: todoId,
        message_link: todo.message_link,
        timestamp: new Date().toISOString(),
        source: 'oracle-chrome-extension-sidepanel'
      }))
    });
    if (!res.ok) throw new Error('Failed to fetch conversation');
    const responseText = await res.text();
    let data = {};
    if (responseText && responseText.trim()) { try { data = JSON.parse(responseText); } catch (e) { data = {}; } }

    // Parse transcript messages (matching newtab parsing exactly)
    let messages = [];
    const responseData = Array.isArray(data) ? data[0] : data;
    const transcript = responseData?.transcript || [];

    // Extract participants for recipients pre-population (matching newtab)
    const participants = responseData?.participants || [];
    const toParticipants = responseData?.to_participants || [];
    const ccParticipants = responseData?.cc_participants || [];
    overlay._participants = participants;
    overlay._toParticipants = toParticipants;
    overlay._ccParticipants = ccParticipants;

    // Detect email task and show recipients button
    const isEmailTask = participants.length > 0;
    const recipientsBtn = overlay.querySelector('.transcript-recipients-btn');
    if (isEmailTask && recipientsBtn) recipientsBtn.style.display = 'flex';

    if (transcript && transcript.length > 0) {
      try {
        const flatTranscript = transcript.flat();
        messages = flatTranscript.map(msgStr => {
          try { return typeof msgStr === 'string' ? JSON.parse(msgStr) : msgStr; } catch { return null; }
        }).filter(m => m !== null);
      } catch (e) { console.error('Error parsing transcript:', e); }
    }

    if (countEl) countEl.textContent = messages.length > 0 ? `${messages.length} message${messages.length !== 1 ? 's' : ''}` : 'No messages';
    // Store messages on overlay for Oracle access
    overlay.transcriptMessages = messages;

    if (messages.length > 0) {
      messagesEl.innerHTML = '';
      messages.forEach((msg, idx) => {
        const displayTime = msg.time ? formatDate(msg.time) : '';
        const senderInitial = (msg.message_from || 'U').charAt(0).toUpperCase();
        const senderName = msg.message_from || 'Unknown';
        const msgContent = msg.message || '';
        const htmlContent = msg.message_html || msgContent;

        const msgDiv = document.createElement('div');
        msgDiv.style.cssText = 'display:flex;flex-direction:column;gap:6px;';

        // Sender header with avatar
        const headerDiv = document.createElement('div');
        headerDiv.style.cssText = 'display:flex;align-items:center;gap:8px;';
        headerDiv.innerHTML = `
          <div style="width:28px;height:28px;background:linear-gradient(135deg,#667eea,#764ba2);border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:11px;font-weight:600;flex-shrink:0;">${senderInitial}</div>
          <div style="flex:1;min-width:0;">
            <div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">${escapeHtml(senderName)}</div>
            <div style="font-size:10px;color:${isDark ? '#666' : '#95a5a6'};">${displayTime}</div>
          </div>`;

        // Message bubble
        const contentDiv = document.createElement('div');
        contentDiv.style.cssText = `margin-left:36px;padding:10px 14px;background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border-radius:12px;border-top-left-radius:4px;font-size:13px;color:${isDark ? '#e0e0e0' : '#2c3e50'};line-height:1.6;word-wrap:break-word;overflow-wrap:break-word;word-break:break-word;overflow-x:hidden;max-width:100%;`;
        if (msg.message_html) {
          renderEmailInIframe(msg.message_html, contentDiv, isDark);
        } else if (isComplexEmailHtml(msgContent)) {
          renderEmailInIframe(msgContent, contentDiv, isDark);
        } else {
          contentDiv.innerHTML = formatMessageContent(msgContent);
        }

        msgDiv.appendChild(headerDiv);
        msgDiv.appendChild(contentDiv);
        messagesEl.appendChild(msgDiv);
      });
      messagesEl.scrollTop = messagesEl.scrollHeight;
    } else {
      // No transcript - show task details as fallback
      messagesEl.innerHTML = `<div style="padding:12px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-radius:12px;">
        <div style="font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};line-height:1.6;white-space:pre-wrap;">${escapeHtml(todo.task_name || '')}</div>
        ${todo.due_by ? `<div style="margin-top:8px;"><span class="todo-due ${new Date(todo.due_by) < new Date() ? 'overdue' : ''}" style="font-size:11px;padding:3px 8px;">${formatDueBy(todo.due_by)}</span></div>` : ''}
      </div>`;
    }
  } catch (e) {
    console.error('Transcript fetch error:', e);
    messagesEl.innerHTML = `<div style="padding:12px;background:${isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'};border-radius:12px;">
      <div style="font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};line-height:1.6;white-space:pre-wrap;">${escapeHtml(todo.task_name || '')}</div>
    </div>`;
    if (countEl) countEl.textContent = 'Task details';
  }
}

// ==================== CHAT & NEW MESSAGE — via shared components ====================
function showChatSlider() {
  const area = document.getElementById('mainContentArea');
  window.OracleAssistant.showChatSlider({ mode: 'inline', container: area });
}
function showNewMessageSlider() {
  const area = document.getElementById('mainContentArea');
  window.OracleNewMessage.showNewMessageSlider({ mode: 'inline', container: area, source: 'oracle-sidepanel' });
}

// ==================== PENDING UPDATES ====================
function showPendingUpdatesBanner(count) {
  const container = document.getElementById('todosContainer'); if (!container) return;
  let banner = container.querySelector('.pending-updates-banner');
  if (banner) { const c = banner.querySelector('.pending-count'); if (c) c.textContent = count; return; }
  container.insertAdjacentHTML('afterbegin', `<div class="pending-updates-banner"><div class="pending-updates-content"><span class="pending-icon">↑</span><span class="pending-text"><span class="pending-count">${count}</span> new update${count > 1 ? 's' : ''}</span></div></div>`);
  banner = container.querySelector('.pending-updates-banner');
  if (banner) banner.addEventListener('click', applyPendingUpdates);
}
function applyPendingUpdates() {
  if (!pendingActionData) return;
  allTodos = pendingActionData; pendingActionData = null; pendingActionUpdates = [];
  allTodos.forEach(t => { const s = String(t.id); if (t.updated_at) { const prev = previousTaskTimestamps.get(s); if (prev && prev !== t.updated_at) readTaskIds.delete(s); previousTaskTimestamps.set(s, t.updated_at); } });
  saveReadState(); displayActions(allTodos);
}

// ==================== DARK MODE ====================
function initDarkMode() { if (localStorage.getItem('oracle_dark_mode') === 'true') document.body.classList.add('dark-mode'); updateDarkModeIcon(); }
function toggleDarkMode() { document.body.classList.toggle('dark-mode'); localStorage.setItem('oracle_dark_mode', document.body.classList.contains('dark-mode')); updateDarkModeIcon(); }
function updateDarkModeIcon() { const b = document.getElementById('darkModeToggle'); if (b) b.textContent = document.body.classList.contains('dark-mode') ? '🌙' : '☀️'; }

// ==================== MAIN INTERFACE ====================
function initializeMainInterface() {
  const logoutBtn = document.getElementById('logoutBtn');
  if (logoutBtn) { logoutBtn.style.display = 'inline-block'; logoutBtn.addEventListener('click', async () => { await logout(); location.reload(); }); }
  document.getElementById('darkModeToggle')?.addEventListener('click', toggleDarkMode);
  document.getElementById('chatToggle')?.addEventListener('click', showChatSlider);
  document.getElementById('newMessageToggle')?.addEventListener('click', showNewMessageSlider);
  loadTodos();
}

// ==================== MESSAGE LISTENERS ====================
try { chrome.runtime.onMessage.addListener((request) => { if (request.action === 'todoListUpdated') performSilentFetch(); }); } catch (e) { }

async function performSilentFetch() {
  try {
    const res = await fetch(WEBHOOK_URL, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(createAuthenticatedPayload({ action: 'list_todos', filter: 'all', timestamp: new Date().toISOString() })) });
    if (!res.ok) return;
    const text = await res.text(); let data = []; if (text?.trim()) try { data = JSON.parse(text); } catch (e) { }
    const newTodos = Array.isArray(data) ? data : (data?.todos || []);
    const activeTodos = newTodos.filter(t => t.status === 0 && !isRecentlyCompleted(t.id));
    const currentIds = new Set(allTodos.filter(t => !isMeetingLink(t.message_link)).map(t => t.id));
    const newItems = activeTodos.filter(t => !isMeetingLink(t.message_link) && !currentIds.has(t.id));
    const updatedItems = activeTodos.filter(t => { if (!currentIds.has(t.id)) return false; const prev = previousTaskTimestamps.get(String(t.id)); return prev && t.updated_at && prev !== t.updated_at; });
    const total = newItems.length + updatedItems.length;
    if (total > 0) { pendingActionData = activeTodos; const ids = [...newItems.map(t => t.id), ...updatedItems.map(t => t.id)]; const existing = new Set(pendingActionUpdates.map(String)); ids.forEach(id => existing.add(String(id))); pendingActionUpdates = Array.from(existing); showPendingUpdatesBanner(pendingActionUpdates.length); }
  } catch (e) { console.error('Silent fetch error:', e); }
}

// ==================== INIT ====================
document.addEventListener('DOMContentLoaded', async () => {
  initDarkMode(); loadReadState();
  // Sync local variables from shared state
  const _st = window.Oracle.state;
  readTaskIds = _st.readTaskIds;
  previousTaskTimestamps = _st.previousTaskTimestamps;
  isInitialLoad = _st.isInitialLoad;
  if (await initAuth()) initializeMainInterface(); else showLoginScreen();
});
