// background.js - Updated with Ably push notifications
const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';
const STORAGE_KEY = 'oracle_user_data';
const ABLY_API_KEY = 'IROXlg.Pr4FZw:5pU-xl09axAY1_jHcnegOd6aJQBXXCiCfAjVXOAQzZI';

// Ably WebSocket connection (native implementation - no SDK needed)
let ablyWebSocket = null;
let currentUserId = null;

// ============================================
// KEEP-ALIVE MECHANISM FOR SERVICE WORKER
// ============================================
// Chrome MV3 service workers can go inactive after ~30 seconds
// This keeps the service worker alive to maintain Ably connection

const KEEP_ALIVE_INTERVAL = 20; // seconds (must be < 30)

// Create keep-alive alarm on startup
function setupKeepAlive() {
  // Create alarm that fires every 20 seconds
  chrome.alarms.create('keepAlive', { 
    periodInMinutes: KEEP_ALIVE_INTERVAL / 60 
  });
  console.log('⏰ Keep-alive alarm set up');
}

// Handle keep-alive alarm
chrome.alarms.onAlarm.addListener(async (alarm) => {
  if (alarm.name === 'keepAlive') {
    // This keeps the service worker alive
    // Also check if Ably connection is still active
    const isConnected = ablyWebSocket && ablyWebSocket.readyState === EventSource.OPEN;
    
    if (!isConnected) {
      console.log('🔄 Keep-alive: Ably connection not open, reconnecting...');
      // Close stale connection if exists
      if (ablyWebSocket) {
        try {
          ablyWebSocket.close();
        } catch (e) {}
        ablyWebSocket = null;
      }
      await initializeAbly();
    } else {
      // Just a ping to keep alive
      console.log('💓 Keep-alive ping, Ably connected');
    }
    return;
  }
  
  if (alarm.name === 'periodicCheck') {
    // Single periodic check every 10 minutes
    checkAndScheduleNotifications();
    return;
  }
});

// Handle extension icon click to open side panel
chrome.action.onClicked.addListener((tab) => {
  chrome.sidePanel.open({ windowId: tab.windowId });
});

// Helper function to safely get domain from URL
function getDomainSafe(url) {
  try {
    return new URL(url).hostname;
  } catch (error) {
    if (url.startsWith('chrome://')) {
      return 'chrome-extension';
    } else if (url.startsWith('file://')) {
      return 'local-file';
    } else {
      return 'unknown-domain';
    }
  }
}

// Get authenticated user data
async function getAuthenticatedUserData() {
  try {
    const result = await chrome.storage.local.get([STORAGE_KEY]);
    if (result[STORAGE_KEY] && result[STORAGE_KEY].userId) {
      return result[STORAGE_KEY];
    }
    return null;
  } catch (error) {
    console.error('Error getting user data:', error);
    return null;
  }
}

// Create authenticated payload
function createAuthenticatedPayload(basePayload, userData) {
  return {
    ...basePayload,
    user_id: userData.userId,
    authenticated: true
  };
}

// Check for upcoming due tasks and schedule notifications
async function checkAndScheduleNotifications() {
  try {
    const userData = await getAuthenticatedUserData();
    if (!userData) return;

    const basePayload = {
      action: 'list_todos',
      filter: 'all',
      timestamp: new Date().toISOString()
    };

    const authenticatedPayload = createAuthenticatedPayload(basePayload, userData);

    const response = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(authenticatedPayload)
    });

    if (response.ok) {
      const data = await response.json();
      const todos = Array.isArray(data) ? data : (data.todos || []);

      // Update badge based on overdue/starred count
      const now = new Date();
      const overdueTodos = todos.filter(todo =>
        todo.status === 0 &&
        todo.due_by &&
        new Date(todo.due_by) < now
      );

      const overdueCount = overdueTodos.length;

      if (overdueCount > 0) {
        chrome.action.setBadgeText({ text: overdueCount.toString() });
        chrome.action.setBadgeBackgroundColor({ color: "#e74c3c" });
      } else {
        const starredCount = todos.filter(todo => todo.starred === 1 || todo.starred === true).length;

        if (starredCount > 0) {
          chrome.action.setBadgeText({ text: starredCount.toString() });
          chrome.action.setBadgeBackgroundColor({ color: "#667eea" });
        } else {
          chrome.action.setBadgeText({ text: "" });
        }
      }
    }
  } catch (error) {
    console.error('Error checking todos:', error);
  }
}

// Handle notification clicks
chrome.notifications.onClicked.addListener((notificationId) => {
  chrome.action.openPopup();
  chrome.notifications.clear(notificationId);
});

chrome.notifications.onButtonClicked.addListener((notificationId, buttonIndex) => {
  if (buttonIndex === 0) {
    chrome.action.openPopup();
    chrome.notifications.clear(notificationId);
  }
});

// Handle messages from content scripts
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'refreshTodoList') {
    // Refresh badge and notify any open panels to refresh
    checkAndScheduleNotifications().then(() => {
      // Broadcast refresh message to all tabs (for side panel)
      chrome.runtime.sendMessage({ action: 'todoListUpdated' }).catch(() => {
        // Ignore errors if no listeners
      });
      sendResponse({ success: true });
    });
    return true; // Keep channel open for async response
  }

  if (request.action === 'makeWebhookCall') {
    // Handle webhook calls from content script (avoids CORS issues)
    const WEBHOOK_URL = 'https://n8n-kqq5.onrender.com/webhook/d60909cd-7ae4-431e-9c4f-0c9cacfa20ea';

    fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(request.payload)
    })
      .then(response => {
        if (response.ok) {
          return response.json().catch(() => ({}));
        } else {
          throw new Error(`HTTP ${response.status}: ${response.statusText}`);
        }
      })
      .then(data => {
        sendResponse({ success: true, data: data });
      })
      .catch(error => {
        console.error('Webhook call error:', error);
        sendResponse({ success: false, error: error.message });
      });

    return true; // Keep channel open for async response
  }
});

// Create context menu when extension installs
chrome.runtime.onInstalled.addListener(() => {
  // Create main Oracle menu item
  chrome.contextMenus.create({
    id: "oracle-main",
    title: "Oracle",
    contexts: ["all"]
  });

  chrome.contextMenus.create({
    id: "oracle-learn",
    title: "Learn",
    contexts: ["all"],
    parentId: "oracle-main"
  });

  chrome.contextMenus.create({
    id: "oracle-task",
    title: "Task",
    contexts: ["all"],
    parentId: "oracle-main"
  });

  chrome.contextMenus.create({
    id: "oracle-priority-task",
    title: "Priority Task",
    contexts: ["all"],
    parentId: "oracle-main"
  });

  chrome.contextMenus.create({
    id: "oracle-bookmark",
    title: "Bookmark",
    contexts: ["all"],
    parentId: "oracle-main"
  });

  chrome.contextMenus.create({
    id: "oracle-fd-ticket-summary",
    title: "FD Ticket Summary",
    contexts: ["all"],
    parentId: "oracle-main"
  });

  // Initial updates
  checkAndScheduleNotifications();
  
  // Set up keep-alive for Ably connection
  setupKeepAlive();
  
  // Initialize Ably connection
  initializeAbly();
});

// Update badge when extension starts
chrome.runtime.onStartup.addListener(() => {
  console.log('🚀 Extension starting up...');
  checkAndScheduleNotifications();
  
  // Set up keep-alive for Ably connection
  setupKeepAlive();
  
  // Initialize Ably connection
  initializeAbly();
});

// Periodic check - DISABLED (using Ably push notifications instead)
// Uncomment this line if you want a backup check every 60 minutes
// chrome.alarms.create('periodicCheck', { periodInMinutes: 60 });

// Handle context menu clicks with authentication
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  console.log("Context menu clicked:", info.menuItemId);

  const userData = await getAuthenticatedUserData();
  if (!userData) {
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icon.png',
      title: 'Oracle - Authentication Required',
      message: 'Please open Oracle popup or new tab to log in first.',
      requireInteraction: true
    });

    chrome.action.setBadgeText({ text: "!" });
    chrome.action.setBadgeBackgroundColor({ color: "#f44336" });

    setTimeout(() => {
      chrome.action.setBadgeText({ text: "" });
    }, 3000);

    return;
  }

  let basePayload = {
    timestamp: new Date().toISOString(),
    source: 'oracle-chrome-extension',
    url: tab.url,
    pageTitle: tab.title,
    selectedText: info.selectionText || '',
    menuOption: info.menuItemId,
    userAgent: 'Oracle Chrome Extension',
    domain: getDomainSafe(tab.url),
    fullUrl: info.pageUrl || tab.url
  };

  if (info.menuItemId === 'oracle-learn') {
    basePayload.action = 'learn';
    basePayload.intent = 'learning';
    basePayload.message = `Learning request: "${info.selectionText || 'Page content'}" from ${tab.title}`;
  } else if (info.menuItemId === 'oracle-task') {
    basePayload.action = 'create_task';
    basePayload.intent = 'task_creation';
    basePayload.taskText = info.selectionText || tab.title;
    basePayload.sourceUrl = tab.url;
    basePayload.priority = false;
    basePayload.message = `Create task: "${info.selectionText || tab.title}" from ${tab.title}`;
  } else if (info.menuItemId === 'oracle-priority-task') {
    basePayload.action = 'priority_browser_task';
    basePayload.intent = 'priority_task_modal';
    basePayload.taskText = info.selectionText || tab.title;
    basePayload.sourceUrl = tab.url;
    basePayload.priority = true;
    basePayload.message = `Show priority task modal: "${info.selectionText || tab.title}" from ${tab.title}`;
  } else if (info.menuItemId === 'oracle-bookmark') {
    basePayload.action = 'create_bookmark_modal';
    basePayload.intent = 'bookmark_modal';
    basePayload.bookmarkText = info.selectionText || tab.title;
    basePayload.bookmarkUrl = tab.url;
    basePayload.pageTitle = tab.title;
    basePayload.message = `Show bookmark modal: "${info.selectionText || tab.title}" from ${tab.title}`;
  } else if (info.menuItemId === 'oracle-fd-ticket-summary') {
    basePayload.action = 'fd_ticket_summary';
    basePayload.intent = 'fd_ticket_summary';
    basePayload.selectedText = info.selectionText || '';
    basePayload.message = `FD Ticket Summary request: "${info.selectionText || 'Page content'}" from ${tab.title}`;
  }

  if (info.selectionText) {
    basePayload.hasSelection = true;
    basePayload.selectionLength = info.selectionText.length;
    basePayload.contextType = 'text_selection';
  } else {
    basePayload.hasSelection = false;
    basePayload.contextType = 'page_context';
  }

  const authenticatedPayload = createAuthenticatedPayload(basePayload, userData);

  try {
    chrome.action.setBadgeText({ text: "..." });
    chrome.action.setBadgeBackgroundColor({ color: "#4CAF50" });

    console.log('Sending authenticated payload:', authenticatedPayload);

    const response = await fetch(WEBHOOK_URL, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(authenticatedPayload)
    });

    if (response.ok) {
      let workflowResponse;
      try {
        const responseText = await response.text();
        console.log('Raw response:', responseText);

        // Check if response is empty
        if (!responseText || responseText.trim() === '') {
          throw new Error('Empty response from server');
        }

        workflowResponse = JSON.parse(responseText);
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        throw new Error('Invalid response format from server');
      }

      console.log('Oracle workflow response:', workflowResponse);

      chrome.action.setBadgeText({ text: "✓" });
      chrome.action.setBadgeBackgroundColor({ color: "#4CAF50" });

      // Handle Priority Task Modal (show_modal action)
      if (info.menuItemId === 'oracle-priority-task') {
        try {
          console.log('Attempting to send Priority Task modal...');
          console.log('Modal data:', workflowResponse);
          console.log('Selected text:', info.selectionText || '');
          console.log('Tab ID:', tab.id);

          await chrome.tabs.sendMessage(tab.id, {
            action: 'showPriorityTaskModal',
            data: workflowResponse, // Send full response, let content.js extract
            selectedText: info.selectionText || '',
            sourceUrl: tab.url
          });

          console.log('Modal message sent successfully!');

          // Don't refresh here - will refresh after user submits the modal

          return; // Exit early for modal, don't show notifications
        } catch (e) {
          console.error('ERROR: Could not send modal to content script:', e);
          // Fall through to error handling
        }
      }

      // Handle Bookmark Modal (show_bookmark_modal action)
      if (info.menuItemId === 'oracle-bookmark') {
        try {
          console.log('Attempting to send Bookmark modal...');
          console.log('Modal data:', workflowResponse);
          console.log('Page title:', tab.title);
          console.log('Tab ID:', tab.id);

          await chrome.tabs.sendMessage(tab.id, {
            action: 'showBookmarkModal',
            data: workflowResponse,
            bookmarkUrl: tab.url,
            pageTitle: tab.title,
            selectedText: info.selectionText || ''
          });

          console.log('Bookmark modal message sent successfully!');

          // Don't refresh here - will refresh after user submits the modal

          return; // Exit early for modal, don't show notifications
        } catch (e) {
          console.error('ERROR: Could not send bookmark modal to content script:', e);
          // Fall through to error handling
        }
      }

      let responseText = 'Processing your request...';

      if (Array.isArray(workflowResponse) && workflowResponse.length > 0) {
        const executionData = workflowResponse[0];

        if (executionData.message) {
          responseText = executionData.message;
        } else if (executionData.output) {
          // Handle FD Ticket Summary output
          responseText = executionData.output;
        } else if (executionData.body && executionData.body.selectedText) {
          if (info.menuItemId === 'oracle-task') {
            responseText = `Task created: "${executionData.body.selectedText}"`;
          } else if (info.menuItemId === 'oracle-priority-task') {
            responseText = `Priority task created: "${executionData.body.selectedText}"`;
          } else if (info.menuItemId === 'oracle-bookmark') {
            responseText = `Bookmark created: "${executionData.body.selectedText}"`;
          } else {
            responseText = `Learning: "${executionData.body.selectedText}"`;
          }
        } else {
          if (info.menuItemId === 'oracle-task') {
            responseText = `Oracle created your task`;
          } else if (info.menuItemId === 'oracle-priority-task') {
            responseText = `Oracle created your priority task`;
          } else if (info.menuItemId === 'oracle-bookmark') {
            responseText = `Oracle created your bookmark`;
          } else if (info.menuItemId === 'oracle-fd-ticket-summary') {
            responseText = `Oracle is generating FD Ticket Summary`;
          } else {
            responseText = `Oracle is processing your ${authenticatedPayload.action} request`;
          }
        }
      } else if (workflowResponse && workflowResponse.message) {
        responseText = workflowResponse.message;
      } else if (workflowResponse && workflowResponse.output) {
        // Direct output from workflow
        responseText = workflowResponse.output;
      }

      let actionText = 'Process';
      if (info.menuItemId === 'oracle-learn') actionText = 'Learn';
      else if (info.menuItemId === 'oracle-task') actionText = 'Task';
      else if (info.menuItemId === 'oracle-priority-task') actionText = 'Priority Task';
      else if (info.menuItemId === 'oracle-bookmark') actionText = 'Bookmark';
      else if (info.menuItemId === 'oracle-fd-ticket-summary') actionText = 'FD Ticket Summary';

      // Only show Chrome notification for non-FD Ticket Summary actions
      if (info.menuItemId !== 'oracle-fd-ticket-summary') {
        chrome.notifications.create({
          type: 'basic',
          iconUrl: 'icon.png',
          title: `Oracle - ${actionText}`,
          message: responseText.substring(0, 200) + (responseText.length > 200 ? '...' : ''),
          requireInteraction: true
        });
      }

      // Show alert for FD Ticket Summary only
      if (info.menuItemId === 'oracle-fd-ticket-summary') {
        try {
          await chrome.tabs.sendMessage(tab.id, {
            action: 'showAlert',
            message: responseText,
            title: 'FD Ticket Summary'
          });
        } catch (e) {
          console.log('Could not send alert to content script:', e);
        }
      } else {
        // Show toast for other actions (not FD Ticket Summary, not Priority Task)
        try {
          await chrome.tabs.sendMessage(tab.id, {
            action: 'showToast',
            message: responseText.substring(0, 100) + (responseText.length > 100 ? '...' : ''),
            type: 'success',
            title: `Oracle - ${actionText}`
          });
        } catch (e) {
          console.log('Could not send message to content script:', e);
        }
      }

      setTimeout(() => {
        checkAndScheduleNotifications();
      }, 3000);

    } else {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }

  } catch (error) {
    console.error('Error triggering Oracle workflow:', error);

    chrome.action.setBadgeText({ text: "✗" });
    chrome.action.setBadgeBackgroundColor({ color: "#f44336" });

    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icon.png',
      title: 'Oracle Workflow Error',
      message: `Error: ${error.message}`
    });

    setTimeout(() => {
      checkAndScheduleNotifications();
    }, 3000);
  }
});

// ============================================
// ABLY REAL-TIME PUSH NOTIFICATIONS (Native WebSocket)
// ============================================

// Variables for Ably connection (already declared at top of file)
let reconnectAttempts = 0;
const MAX_RECONNECT_DELAY = 30000;

async function initializeAbly() {
  try {
    const userData = await getAuthenticatedUserData();
    if (!userData || !userData.userId) {
      console.log('User not authenticated, skipping Ably initialization');
      return;
    }

    const userId = userData.userId;
    currentUserId = userId; // Store for reconnection
    console.log(`🔄 Connecting to Ably for user ${userId}...`);

    // Connect to Ably using native WebSocket
    connectAblyWebSocket(userId);

  } catch (error) {
    console.error('❌ Failed to initialize Ably:', error);
  }
}

function connectAblyWebSocket(userId) {
  // Use EventSource (SSE) instead of WebSocket for better stability
  const channel = `user_${userId}`;
  const sseUrl = `https://realtime.ably.io/sse?key=${encodeURIComponent(ABLY_API_KEY)}&channels=${channel}&v=1.2`;

  console.log(`🔄 Connecting to Ably SSE for channel: ${channel}`);

  // Close existing connection
  if (ablyWebSocket) {
    try {
      ablyWebSocket.close();
    } catch (e) {
      console.log('Error closing existing connection:', e);
    }
    ablyWebSocket = null;
  }

  try {
    ablyWebSocket = new EventSource(sseUrl);

    ablyWebSocket.onopen = () => {
      console.log('✅ Ably SSE connected');
      reconnectAttempts = 0;
    };

    ablyWebSocket.onmessage = async (event) => {
      try {
        const data = JSON.parse(event.data);
        console.log('📩 Ably SSE message:', data);

        // Ably SSE sends individual messages, not an array
        if (data.name) {
          // Parse the data field if it's a string
          let messageData = data.data;
          if (typeof messageData === 'string') {
            try {
              messageData = JSON.parse(messageData);
            } catch (e) {
              console.warn('Could not parse message data:', messageData);
            }
          }

          console.log('📨 Received:', data.name, messageData);

          // Handle different event types
          if (data.name === 'todo_created') {
            await handleTodoCreated(messageData);
          } else if (data.name === 'todo_updated') {
            await handleTodoUpdated(messageData);
          } else if (data.name === 'todo_deleted') {
            await handleTodoDeleted(messageData);
          } else if (data.name === 'bookmark_created') {
            await handleBookmarkCreated(messageData);
          }
        }
      } catch (error) {
        console.error('Error processing SSE message:', error, event.data);
      }
    };

    ablyWebSocket.onerror = (event) => {
      // EventSource onerror receives an Event object, not an Error
      // Check readyState to determine the actual issue
      const readyState = ablyWebSocket ? ablyWebSocket.readyState : 'unknown';
      const stateNames = { 0: 'CONNECTING', 1: 'OPEN', 2: 'CLOSED' };
      const stateName = stateNames[readyState] || readyState;
      
      // Only log as warning since keep-alive will handle reconnection
      console.warn(`⚠️ Ably SSE connection issue (state: ${stateName})`);
      
      // If connection is closed, let keep-alive handle reconnection
      // Don't schedule reconnect here - prevents multiple attempts
    };

  } catch (error) {
    console.error('❌ Failed to create SSE connection:', error);
  }
}

// Event handlers
async function handleTodoCreated(data) {
  console.log('📨 New todo created:', data);

  // Refresh todos in background
  await checkAndScheduleNotifications();

  // Notify any open sidepanels or newtabs to refresh
  try {
    await chrome.runtime.sendMessage({
      action: 'todoListUpdated',
      todo: data
    });
    console.log('✅ Sidepanel/Newtab notified to refresh');
  } catch (error) {
    console.log('ℹ️ No sidepanel/newtab open (this is normal)');
  }
}

async function handleTodoUpdated(data) {
  console.log('📝 Todo updated:', data);
  await checkAndScheduleNotifications();

  try {
    await chrome.runtime.sendMessage({ action: 'todoListUpdated', todo: data });
  } catch (error) {
    console.log('ℹ️ No sidepanel/newtab open');
  }
}

async function handleTodoDeleted(data) {
  console.log('🗑️ Todo deleted:', data);
  await checkAndScheduleNotifications();

  try {
    await chrome.runtime.sendMessage({ action: 'todoListUpdated', todoId: data.todoId });
  } catch (error) {
    console.log('ℹ️ No sidepanel/newtab open');
  }
}

async function handleBookmarkCreated(data) {
  console.log('🔖 Bookmark created:', data);

  try {
    await chrome.runtime.sendMessage({ action: 'bookmarksUpdated', bookmark: data });
    console.log('✅ Sidepanel/Newtab notified of bookmark creation');
  } catch (error) {
    console.log('ℹ️ No sidepanel/newtab open');
  }
}

// Initialize on script load (service worker wake)
console.log('🚀 Service worker loaded, setting up...');

// Set up keep-alive immediately
setupKeepAlive();

// Initialize Ably connection
initializeAbly();

// Re-initialize if user logs in/out
chrome.storage.onChanged.addListener((changes, area) => {
  if (area === 'local' && changes[STORAGE_KEY]) {
    console.log('🔄 User authentication changed, reinitializing Ably');

    // Disconnect old connection
    if (ablyWebSocket) {
      try {
        ablyWebSocket.close();
      } catch (e) {}
      ablyWebSocket = null;
    }

    // Initialize new connection
    initializeAbly();
  }
});

// Handle messages that might wake up the service worker
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  // Ensure Ably is connected when service worker wakes up
  if (!ablyWebSocket || ablyWebSocket.readyState !== EventSource.OPEN) {
    console.log('🔄 Message received, checking Ably connection...');
    initializeAbly();
  }
  
  // Continue with normal message handling (handled earlier in file)
  return false;
});