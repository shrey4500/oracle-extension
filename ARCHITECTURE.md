# Oracle Chrome Extension — Architecture

## File Structure

```
Oracle V36/
├── ably.min.js                  # Ably realtime library
├── background.js                # Service worker
├── content.js                   # Content script
├── manifest.json                # Extension manifest
├── newtab.html                  # New tab page
├── newtab.js                    # New tab logic (9,400 lines)
├── sidepanel.html               # Side panel page
├── sidepanel.js                 # Side panel logic (1,350 lines)
├── shared/                      # Shared components (7 files, ~2,600 lines)
│   ├── oracle-common.js         # State, auth, utilities
│   ├── oracle-icons.js          # Link detection, icon mapping
│   ├── oracle-grouping.js       # Task grouping (Drive/Slack/Tag)
│   ├── oracle-message-format.js # Message formatting, HTML sanitization, attachments
│   ├── oracle-assistant.js      # Chat slider (Ably streaming + inline)
│   ├── oracle-new-message.js    # Compose slider (Slack/Gmail)
│   └── oracle-notes.js          # Notes CRUD, All Tasks
└── icon-*.png                   # Platform icons (14 files)
```

## Script Loading Order

Both HTML files load shared components before page-specific scripts:

```
ably.min.js              (newtab only)
shared/oracle-common.js
shared/oracle-icons.js
shared/oracle-grouping.js
shared/oracle-message-format.js
shared/oracle-assistant.js
shared/oracle-new-message.js
shared/oracle-notes.js
newtab.js / sidepanel.js
```

## Shared Component API

### oracle-common.js → `window.Oracle`

State, auth helpers, and utility functions shared across all components.

```
Oracle.state                        — Shared mutable state object
Oracle.WEBHOOK_URL                  — Main webhook endpoint
Oracle.CHAT_WEBHOOK_URL             — Chat webhook endpoint
Oracle.escapeHtml(text)             — HTML escape
Oracle.formatDate(dateString)       — Relative time ("2m ago", "3d ago")
Oracle.formatDueBy(dueDateString)   — Due date display ("Due 3h 20m")
Oracle.sortTodos(todos)             — Sort: due_by → updated_at → created_at
Oracle.isValidTag(tag)              — Filter junk tags
Oracle.createAuthenticatedPayload() — Wraps data with user_id + auth
Oracle.loadReadState()              — Load read/unread tracking from localStorage
Oracle.saveReadState()              — Persist read state
Oracle.markTaskAsRead(id)           — Mark single task read
Oracle.isTaskUnread(id)             — Check unread status
```

### oracle-icons.js → `window.OracleIcons`

Link type detection and platform icon rendering.

```
OracleIcons.isMeetingLink(url)           — Zoom/Meet/Teams/Calendar
OracleIcons.isDriveLink(url)             — Docs/Sheets/Slides/Drive
OracleIcons.isSlackLink(url)             — Slack links
OracleIcons.getSlackChannelUrl(url)      — Extract channel URL
OracleIcons.extractDriveFileId(url)      — Extract file ID from Drive URL
OracleIcons.getCleanDriveFileUrl(url)    — Normalize Drive URL
OracleIcons.getIconForLink(url)          — Returns { icon, bg, label }
OracleIcons.buildSecondaryLinksHtml()    — Render secondary link icons
```

### oracle-grouping.js → `window.OracleGrouping`

Pure logic for grouping tasks by source.

```
OracleGrouping.groupTasksByDriveFile(tasks)    → { driveGroups, nonDriveTasks }
OracleGrouping.groupTasksBySlackChannel(tasks) → { slackGroups, nonSlackTasks }
OracleGrouping.groupTasksByTag(tasks)          → { tagGroups, untaggedTasks }
OracleGrouping.extractFileNameFromTask(task)   → string
```

### oracle-message-format.js → `window.OracleMessageFormat`

Message rendering, HTML sanitization, email display, and attachment handling.

```
OracleMessageFormat.formatMessageContent(text)     — Slack/markdown → HTML
OracleMessageFormat.sanitizeHtml(html)             — Strip dangerous elements
OracleMessageFormat.isComplexEmailHtml(html)       — Detect complex email HTML
OracleMessageFormat.renderEmailInIframe(html, el)  — Sandboxed email iframe
OracleMessageFormat.fetchGmailAttachment(...)      — Fetch via webhook
OracleMessageFormat.renderTranscriptAttachment(att)— Build attachment element
OracleMessageFormat.showAttachmentPreview(url,type)— Full-screen preview modal
```

### oracle-assistant.js → `window.OracleAssistant`

Chat slider with two modes.

```
OracleAssistant.showChatSlider({
  mode: 'fullscreen',  // newtab: fixed overlay, Ably streaming
  mode: 'inline',      // sidepanel: in-container, simple fetch
  container: element,  // required for inline mode
  onClose: callback
})
OracleAssistant.formatChatResponseWithAnnotations(text) — URL → platform icons
```

### oracle-new-message.js → `window.OracleNewMessage`

New message composer supporting Slack and Gmail.

```
OracleNewMessage.showNewMessageSlider({
  mode: 'fullscreen' | 'inline',
  container: element,
  source: string,
  onClose: callback
})
```

Features: platform toggle, recipient search with debounce, @mentions,
channel/group DM restrictions, Gmail CC/subject, keyboard navigation.

### oracle-notes.js → `window.OracleNotes`

Notes CRUD and All Tasks (completed tasks) view.

```
OracleNotes.loadNotes()              — Fetch + render notes
OracleNotes.showNoteForm(note?)      — Create/edit form
OracleNotes.hideNoteForm()           — Close form
OracleNotes.saveNote()               — Save to backend
OracleNotes.showNoteViewer(note)     — Read-only slider
OracleNotes.loadAllTasks()           — Fetch completed tasks
OracleNotes.searchAllTasks(query)    — Search completed tasks
OracleNotes.setupAllTasksSearch()    — Wire search handlers
OracleNotes.setupNoteButtons()       — Wire add/save/cancel/refresh buttons
```

## Auth State Sync

newtab.js and sidepanel.js manage their own `isAuthenticated` and `userData`
variables. After successful auth, they sync to `window.Oracle.state` so shared
components can check authentication:

```js
if (window.Oracle && window.Oracle.state) {
  window.Oracle.state.isAuthenticated = true;
  window.Oracle.state.userData = userData;
}
```

## Key Patterns

- **Dual-mode sliders**: `'fullscreen'` (newtab, fixed overlay 450px) vs `'inline'` (sidepanel, inside container)
- **Dark mode**: All components check `document.body.classList.contains('dark-mode')`
- **Event cleanup**: `removeEventListener` on close, 250ms animation timeout
- **IIFE namespacing**: Each shared file wraps in `(function(){ 'use strict'; ... })()` and exports to `window.*`
- **Dependency chain**: Components destructure from `window.Oracle` — common.js must load first

## What Lives Where

| Feature | Location | Why |
|---------|----------|-----|
| Transcript slider (~3K lines) | newtab.js | Tightly coupled to 20+ local DOM builders |
| Actions tab renderer | newtab.js | Uses buildSingleTodoHtml, meetings accordion |
| FYI tab renderer | newtab.js | Same builders as Actions |
| Ably realtime updates | newtab.js | Manages pending updates, badge counts |
| Keyboard navigation | newtab.js | Page-specific DOM traversal |
| Profile management | newtab.js | Standalone section at end of file |
| formatTimeAgoFresh | newtab.js | Transcript-specific time parser (IST) |
