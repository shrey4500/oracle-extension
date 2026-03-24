// oracle-assistant.js — Quick Chat / Oracle assistant slider
// Supports two modes: 'fullscreen' (newtab) and 'inline' (sidepanel)
// Both modes use Ably streaming for real-time responses
// V37: Added @mention support, copy buttons on bot messages

(function () {
  'use strict';

  const { escapeHtml, WEBHOOK_URL, state } = window.Oracle;

  // Inject styles for chat link icons, mention dropdown, and copy buttons
  if (!document.getElementById('oracle-assistant-styles')) {
    const s = document.createElement('style'); s.id = 'oracle-assistant-styles';
    s.textContent = `
      .oracle-chat-link-icon:hover{background:rgba(102,126,234,0.2)!important}
      .chat-mention-dropdown{position:absolute;bottom:100%;left:0;right:0;max-height:200px;overflow-y:auto;border-radius:8px;box-shadow:0 -4px 12px rgba(0,0,0,0.15);z-index:1000;margin-bottom:4px;display:none;background:#fff;border:1px solid #e1e8ed;}
      .chat-mention-dropdown.show{display:block;}
      .chat-mention-item{display:flex;align-items:center;gap:10px;padding:8px 12px;cursor:pointer;transition:background 0.15s;border-bottom:1px solid rgba(225,232,237,0.5);}
      .chat-mention-item:last-child{border-bottom:none;}
      .chat-mention-item:hover,.chat-mention-item.active{background:rgba(102,126,234,0.12);}
      .chat-mention-item .mention-avatar{width:28px;height:28px;border-radius:50%;background:linear-gradient(45deg,#667eea,#764ba2);display:flex;align-items:center;justify-content:center;color:white;font-size:11px;font-weight:600;flex-shrink:0;}
      .chat-mention-item .mention-name{font-weight:600;font-size:13px;color:#2c3e50;}
      .chat-mention-item .mention-email{font-size:11px;color:#7f8c8d;}
      .chat-mention-chip{display:inline;background:rgba(102,126,234,0.15);color:#667eea;padding:1px 4px;border-radius:4px;font-weight:600;font-size:inherit;white-space:nowrap;}
      body.dark-mode .chat-mention-dropdown{background:#1a2332;border:1px solid rgba(255,255,255,0.1);}
      body.dark-mode .chat-mention-item{border-bottom-color:rgba(255,255,255,0.06);}
      body.dark-mode .chat-mention-item:hover,body.dark-mode .chat-mention-item.active{background:rgba(102,126,234,0.2);}
      body.dark-mode .chat-mention-item .mention-name{color:#e8e8e8;}
      body.dark-mode .chat-mention-item .mention-email{color:#888;}
      .bot-msg-actions{display:flex;gap:4px;justify-content:flex-end;margin-top:6px;opacity:0;transition:opacity 0.2s;}
      .bot-msg-wrapper:hover .bot-msg-actions{opacity:1;}
      .bot-action-btn{background:none;border:1px solid rgba(102,126,234,0.25);color:#667eea;width:26px;height:26px;border-radius:6px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:13px;transition:all 0.15s;padding:0;}
      .bot-action-btn:hover{background:rgba(102,126,234,0.1);border-color:rgba(102,126,234,0.5);}
      .bot-action-btn.copied{background:rgba(46,204,113,0.15);border-color:rgba(46,204,113,0.4);color:#27ae60;}
      body.dark-mode .bot-action-btn{border-color:rgba(102,126,234,0.3);color:#8fa4f0;}
      body.dark-mode .bot-action-btn:hover{background:rgba(102,126,234,0.2);}
    `;
    document.head.appendChild(s);
  }

  function getAblyChatKey() {
    return window.ABLY_CHAT_API_KEY || null;
  }

  async function getFreshUserId() {
    const fromState = window.Oracle?.state?.userData?.userId || state.userData?.userId;
    if (fromState) return fromState;
    try {
      const r = await chrome.storage.local.get(['oracle_user_data']);
      return r?.oracle_user_data?.userId || null;
    } catch { return null; }
  }

  // ============================================
  // Extract links from raw text
  // ============================================
  function extractLinks(text) {
    if (!text) return [];
    const urlRegex = /https?:\/\/[^\s\]\)]+/g;
    return [...new Set((text.match(urlRegex) || []))];
  }

  // ============================================
  // Create copy action buttons for bot messages
  // ============================================
  function createBotActions(rawText) {
    const links = extractLinks(rawText);
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'bot-msg-actions';

    const linkSvg = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/></svg>`;
    const copySvg = `<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`;

    function makeCopyBtn(title, svg, textToCopy) {
      const btn = document.createElement('button');
      btn.className = 'bot-action-btn';
      btn.title = title;
      btn.innerHTML = svg;
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        navigator.clipboard.writeText(textToCopy).then(() => {
          btn.classList.add('copied');
          btn.innerHTML = '✓';
          setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = svg; }, 1500);
        });
      });
      return btn;
    }

    // Button 1: Copy links (only if links exist)
    if (links.length > 0) {
      actionsDiv.appendChild(makeCopyBtn('Copy links', linkSvg, links.join('\n')));
    }
    // Button 2: Copy full text (always)
    actionsDiv.appendChild(makeCopyBtn('Copy message', copySvg, rawText));

    return actionsDiv;
  }

  // ============================================
  // formatChatResponseWithAnnotations
  // ============================================
  function formatChatResponseWithAnnotations(text) {
    if (!text) return '';

    let slackIconUrl, gmailIconUrl, driveIconUrl, googleDocsIconUrl, googleSheetsIconUrl,
      googleSlidesIconUrl, freshdeskIconUrl, freshreleaseIconUrl, freshserviceIconUrl;
    try {
      slackIconUrl = chrome.runtime.getURL('icon-slack.png');
      gmailIconUrl = chrome.runtime.getURL('icon-gmail.png');
      driveIconUrl = chrome.runtime.getURL('icon-drive.png');
      googleDocsIconUrl = chrome.runtime.getURL('icon-google-docs.png');
      googleSheetsIconUrl = chrome.runtime.getURL('icon-google-sheets.png');
      googleSlidesIconUrl = chrome.runtime.getURL('icon-google-slides.png');
      freshdeskIconUrl = chrome.runtime.getURL('icon-freshdesk.png');
      freshreleaseIconUrl = chrome.runtime.getURL('icon-freshrelease.png');
      freshserviceIconUrl = chrome.runtime.getURL('icon-freshservice.png');
    } catch (e) { /* fallback */ }

    const linkIconSvg = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

    function getIconForUrl(url) {
      let iconHtml = linkIconSvg, title = 'Open link';
      const img = (src, alt, t) => { iconHtml = `<img src="${src}" alt="${alt}" style="width:14px;height:14px;vertical-align:middle;">`; title = t; };
      if (url.includes('slack.com')) img(slackIconUrl, 'Slack', 'Open in Slack');
      else if (url.includes('mail.google.com')) img(gmailIconUrl, 'Gmail', 'Open in Gmail');
      else if (url.includes('docs.google.com/document')) img(googleDocsIconUrl, 'Docs', 'Open in Google Docs');
      else if (url.includes('docs.google.com/spreadsheets') || url.includes('sheets.google.com')) img(googleSheetsIconUrl, 'Sheets', 'Open in Google Sheets');
      else if (url.includes('docs.google.com/presentation') || url.includes('slides.google.com')) img(googleSlidesIconUrl, 'Slides', 'Open in Google Slides');
      else if (url.includes('drive.google.com')) img(driveIconUrl, 'Drive', 'Open in Google Drive');
      else if (url.includes('freshdesk.com')) img(freshdeskIconUrl, 'Freshdesk', 'Open in Freshdesk');
      else if (url.includes('freshrelease.com')) img(freshreleaseIconUrl, 'Freshrelease', 'Open in Freshrelease');
      else if (url.includes('freshservice.com')) img(freshserviceIconUrl, 'Freshservice', 'Open in Freshservice');
      return `<a href="${url}" target="_blank" title="${title}" class="oracle-chat-link-icon" style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;background:rgba(102,126,234,0.1);border-radius:4px;margin:0 2px;text-decoration:none;vertical-align:middle;transition:all 0.2s;">${iconHtml}</a>`;
    }

    let result = text;
    result = result.replace(/\[?(https?:\/\/[^\s\]\)]+)\]?/g, (match, url) => getIconForUrl(url));
    const parts = result.split(/(<a href="[^"]*"[^>]*>.*?<\/a>)/g);
    result = parts.map(part => part.startsWith('<a href=') ? part : escapeHtml(part)).join('');
    result = result.replace(/\n/g, '<br>');
    result = result.replace(/\s*-\s*(<a href)/g, ' $1');
    result = result.replace(/\s{2,}/g, ' ');
    return result;
  }

  // ============================================
  // @mention search helper
  // ============================================
  let mentionAbortController = null;
  async function searchPeopleForMention(query) {
    if (mentionAbortController) mentionAbortController.abort();
    mentionAbortController = new AbortController();
    try {
      const res = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'search_user_new_slack',
          query,
          platform: 'slack',
          context: 'mention',
          timestamp: new Date().toISOString(),
          source: 'oracle-assistant-mention',
          user_id: state.userData?.userId,
          authenticated: true
        }),
        signal: mentionAbortController.signal
      });
      if (!res.ok) return [];
      let d = await res.json();
      if (!Array.isArray(d)) d = d.results || d.members || d.users || [];
      return d.filter(r => (r.type === 'employee' || !r.type) && (r.user_email_ID || r.email)).slice(0, 5);
    } catch (e) {
      if (e.name !== 'AbortError') console.error('Mention search error:', e);
      return [];
    }
  }

  // ============================================
  // Setup @mention on a contenteditable input
  // ============================================
  function setupMentionSupport(chatInput, dropdownContainer, isDark) {
    const dropdown = document.createElement('div');
    dropdown.className = 'chat-mention-dropdown';
    dropdownContainer.appendChild(dropdown);

    let mentionActive = false;
    let mentionSearchTimeout = null;
    let activeIndex = -1;
    let cachedPeople = [];

    function getTextBeforeCaret() {
      const sel = window.getSelection();
      if (!sel.rangeCount) return '';
      const range = sel.getRangeAt(0).cloneRange();
      range.collapse(true);
      const treeWalker = document.createTreeWalker(chatInput, NodeFilter.SHOW_TEXT, null);
      let fullText = '';
      let node;
      while (node = treeWalker.nextNode()) {
        if (node === range.startContainer) {
          fullText += node.textContent.substring(0, range.startOffset);
          break;
        }
        fullText += node.textContent;
      }
      return fullText;
    }

    function closeMentionDropdown() {
      dropdown.classList.remove('show');
      dropdown.innerHTML = '';
      mentionActive = false;
      activeIndex = -1;
      cachedPeople = [];
    }

    function insertMention(person) {
      const name = person.Full_Name || person['Full Name'] || person.full_name || person.name || 'Unknown';
      const email = person.user_email_ID || person.email || '';

      const sel = window.getSelection();
      if (!sel.rangeCount) return;

      const textBefore = getTextBeforeCaret();
      const atIdx = textBefore.lastIndexOf('@');
      if (atIdx === -1) return;

      const range = sel.getRangeAt(0);
      const startContainer = range.startContainer;
      const startOffset = range.startOffset;

      const treeWalker = document.createTreeWalker(chatInput, NodeFilter.SHOW_TEXT, null);
      let charCount = 0;
      let atNode = null, atNodeOffset = 0;
      let textNode;
      while (textNode = treeWalker.nextNode()) {
        const nodeLen = textNode.textContent.length;
        if (charCount + nodeLen > atIdx) {
          atNode = textNode;
          atNodeOffset = atIdx - charCount;
          break;
        }
        charCount += nodeLen;
      }
      if (!atNode) return;

      const deleteRange = document.createRange();
      deleteRange.setStart(atNode, atNodeOffset);
      deleteRange.setEnd(startContainer, startOffset);
      deleteRange.deleteContents();

      const chip = document.createElement('span');
      chip.className = 'chat-mention-chip';
      chip.contentEditable = 'false';
      chip.dataset.email = email;
      chip.dataset.name = name;
      chip.textContent = `@${name}`;
      chip.title = email;

      const insertRange = sel.getRangeAt(0);
      insertRange.insertNode(chip);
      const space = document.createTextNode('\u00A0');
      chip.parentNode.insertBefore(space, chip.nextSibling);

      const newRange = document.createRange();
      newRange.setStartAfter(space);
      newRange.collapse(true);
      sel.removeAllRanges();
      sel.addRange(newRange);

      closeMentionDropdown();
    }

    function renderMentionResults(people) {
      if (!people || !people.length) { closeMentionDropdown(); return; }
      cachedPeople = people;
      activeIndex = 0;
      dropdown.innerHTML = people.map((p, i) => {
        const name = p.Full_Name || p['Full Name'] || p.full_name || p.name || 'Unknown';
        const email = p.user_email_ID || p.email || '';
        const initials = name.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
        return `<div class="chat-mention-item ${i === 0 ? 'active' : ''}" data-idx="${i}">
          <div class="mention-avatar">${initials}</div>
          <div style="flex:1;min-width:0;">
            <div class="mention-name">${escapeHtml(name)}</div>
            <div class="mention-email">${escapeHtml(email)}</div>
          </div>
        </div>`;
      }).join('');
      dropdown.classList.add('show');

      dropdown.querySelectorAll('.chat-mention-item').forEach((item, idx) => {
        item.addEventListener('mousedown', (e) => {
          e.preventDefault();
          insertMention(people[idx]);
        });
      });
    }

    chatInput.addEventListener('input', () => {
      const textBefore = getTextBeforeCaret();
      const mentionMatch = textBefore.match(/@(\w{0,30})$/);

      if (mentionMatch) {
        mentionActive = true;
        const query = mentionMatch[1];
        if (query.length >= 2) {
          clearTimeout(mentionSearchTimeout);
          mentionSearchTimeout = setTimeout(async () => {
            const results = await searchPeopleForMention(query);
            renderMentionResults(results);
          }, 300);
        } else if (query.length === 0) {
          closeMentionDropdown();
        }
      } else {
        closeMentionDropdown();
      }
    });

    chatInput.addEventListener('keydown', (e) => {
      if (!dropdown.classList.contains('show')) return;
      const items = dropdown.querySelectorAll('.chat-mention-item');
      if (!items.length) return;

      if (e.key === 'ArrowDown') {
        e.preventDefault();
        items[activeIndex]?.classList.remove('active');
        activeIndex = (activeIndex + 1) % items.length;
        items[activeIndex]?.classList.add('active');
        items[activeIndex]?.scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        items[activeIndex]?.classList.remove('active');
        activeIndex = (activeIndex - 1 + items.length) % items.length;
        items[activeIndex]?.classList.add('active');
        items[activeIndex]?.scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter' && mentionActive) {
        e.preventDefault();
        e.stopPropagation();
        if (cachedPeople[activeIndex]) insertMention(cachedPeople[activeIndex]);
      } else if (e.key === 'Escape') {
        e.preventDefault();
        closeMentionDropdown();
      }
    });

    chatInput.addEventListener('blur', () => {
      setTimeout(() => closeMentionDropdown(), 200);
    });

    return { closeMentionDropdown };
  }

  // ============================================
  // Extract message text including mention data
  // Returns { display: "text with @Name chips", api: "text with Name (email)" }
  // Walks ALL nodes recursively to handle nested contenteditable structures
  // ============================================
  function extractMessageWithMentions(chatInput) {
    let apiText = '';
    let displayText = '';

    function walkNodes(parent) {
      parent.childNodes.forEach(node => {
        if (node.nodeType === Node.TEXT_NODE) {
          apiText += node.textContent;
          displayText += node.textContent;
        } else if (node.nodeType === Node.ELEMENT_NODE) {
          if (node.classList?.contains('chat-mention-chip')) {
            const name = node.dataset.name || '';
            const email = node.dataset.email || '';
            apiText += `${name} (${email})`;
            displayText += `@${name}`;
          } else if (node.tagName === 'BR') {
            apiText += '\n';
            displayText += '\n';
          } else if (node.tagName === 'DIV' || node.tagName === 'P') {
            // Contenteditable wraps lines in divs
            if (apiText.length > 0 && !apiText.endsWith('\n')) {
              apiText += '\n';
              displayText += '\n';
            }
            walkNodes(node);
          } else {
            walkNodes(node);
          }
        }
      });
    }

    walkNodes(chatInput);
    return { api: apiText.trim(), display: displayText.trim() };
  }

  // ============================================
  // Ably streaming send logic
  // ============================================
  async function sendWithAblyStreaming({ message, conversationHistory, messagesContainer, loadingDiv, botMsgDiv, botBubble, isDark, source, onStreamDone }) {
    const userId = await getFreshUserId();
    const sessionId = `chat-${userId || 'anon'}-${Date.now()}`;
    let fullResponseText = '', streamStarted = false, streamTimeout = null;

    const ablyKey = getAblyChatKey();
    if (!ablyKey || typeof Ably === 'undefined') throw new Error('Ably not configured');

    const ablyRealtime = new Ably.Realtime({ key: ablyKey });
    const chatChannel = ablyRealtime.channels.get(sessionId);
    await new Promise((resolve, reject) => { chatChannel.attach((err) => err ? reject(err) : resolve()); });

    const streamPromise = new Promise((resolve, reject) => {
      streamTimeout = setTimeout(() => { chatChannel.unsubscribe(); ablyRealtime.close(); streamStarted ? resolve(fullResponseText) : reject(new Error('Response timeout')); }, 120000);

      chatChannel.subscribe('token', (msg) => {
        if (!streamStarted) { streamStarted = true; if (loadingDiv.parentNode) loadingDiv.remove(); messagesContainer.appendChild(botMsgDiv); }
        let d = msg.data; if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = { text: d }; } }
        fullResponseText += (d.text || d);
        try { botBubble.innerHTML = formatChatResponseWithAnnotations(fullResponseText); } catch { botBubble.textContent = fullResponseText; }
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
      });

      chatChannel.subscribe('done', (msg) => {
        clearTimeout(streamTimeout); chatChannel.unsubscribe(); ablyRealtime.close();
        let d = msg.data; if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = {}; } }
        if (d?.fullText) fullResponseText = d.fullText;
        if (!streamStarted) { if (loadingDiv.parentNode) loadingDiv.remove(); messagesContainer.appendChild(botMsgDiv); streamStarted = true; }
        try { botBubble.innerHTML = formatChatResponseWithAnnotations(fullResponseText); } catch { botBubble.textContent = fullResponseText; }
        messagesContainer.scrollTop = messagesContainer.scrollHeight;
        if (onStreamDone) onStreamDone(fullResponseText);
        resolve(fullResponseText);
      });

      chatChannel.subscribe('error', (msg) => {
        clearTimeout(streamTimeout); chatChannel.unsubscribe(); ablyRealtime.close();
        let d = msg.data; if (typeof d === 'string') { try { d = JSON.parse(d); } catch { d = {}; } }
        reject(new Error(d?.message || 'Stream error'));
      });
    });

    fetch(CHAT_WEBHOOK_URL, {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, conversation: conversationHistory, timestamp: new Date().toISOString(), source, session_id: sessionId, user_id: userId })
    }).catch(err => console.error('Webhook fire error:', err));

    return await streamPromise;
  }

  // ============================================
  // showChatSlider — unified for both modes
  // ============================================
  function showChatSlider(opts = {}) {
    const mode = opts.mode || 'fullscreen';
    const container = opts.container || document.body;
    const onClose = opts.onClose || null;
    const isInline = mode === 'inline';
    const source = isInline ? 'oracle-sidepanel-chat' : 'oracle-quick-chat';

    // Remove existing and cleanup
    if (isInline) {
      container.querySelectorAll('.slider-overlay,.chat-slider-overlay').forEach(s => s.remove());
    } else {
      document.querySelectorAll('.chat-slider-overlay').forEach(s => s.remove());
    }

    state.isChatSliderOpen = true;
    const isDark = document.body.classList.contains('dark-mode');

    let overlay, slider, chatInput, sendBtn, messagesContainer;

    if (isInline) {
      // --- INLINE MODE (sidepanel) ---
      slider = document.createElement('div');
      slider.className = 'slider-overlay';
      slider.innerHTML = `
        <div style="padding:14px;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
          <div style="display:flex;align-items:center;gap:10px;">
            <div style="width:32px;height:32px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:16px;color:white;">🤖</div>
            <div>
              <div style="font-weight:600;font-size:14px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Quick Chat</div>
              <div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">Ask me anything</div>
            </div>
          </div>
          <button class="chat-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:30px;height:30px;border-radius:8px;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center;">×</button>
        </div>
        <div class="slider-body" style="display:flex;flex-direction:column;padding:0;">
          <div class="chat-messages" style="flex:1;overflow-y:auto;padding:14px;display:flex;flex-direction:column;gap:10px;">
            <div class="chat-welcome" style="text-align:center;padding:30px 16px;color:${isDark ? '#888' : '#7f8c8d'};"><div style="font-size:36px;margin-bottom:12px;">💬</div><div style="font-size:13px;">Send a message to start chatting</div></div>
          </div>
          <div style="padding:10px 14px;border-top:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};display:flex;gap:8px;flex-shrink:0;position:relative;">
            <div class="chat-mention-container" style="position:relative;flex:1;display:flex;">
              <div class="chat-input" contenteditable="true" placeholder="Ask Oracle... @name to tag" style="flex:1;padding:10px;border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(102,126,234,0.2)'};border-radius:8px;font-size:13px;background:${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.05)'};color:${isDark ? '#e8e8e8' : '#2c3e50'};max-height:100px;overflow-y:auto;outline:none;"></div>
            </div>
            <button class="chat-send-btn" style="padding:10px 14px;background:linear-gradient(45deg,#667eea,#764ba2);color:white;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;flex-shrink:0;">➤</button>
          </div>
          <div style="font-size:10px;color:${isDark ? '#555' : '#b0b0b0'};text-align:center;padding:4px;">Press Enter to send, Shift+Enter for new line · @name to tag people</div>
        </div>`;
      container.appendChild(slider);

      chatInput = slider.querySelector('.chat-input');
      sendBtn = slider.querySelector('.chat-send-btn');
      messagesContainer = slider.querySelector('.chat-messages');

      const closeSlider = () => { slider.classList.add('closing'); state.isChatSliderOpen = false; if (onClose) onClose(); setTimeout(() => slider.remove(), 250); };
      slider.querySelector('.chat-close-btn').addEventListener('click', closeSlider);

      setupMentionSupport(chatInput, slider.querySelector('.chat-mention-container'), isDark);

    } else {
      // --- FULLSCREEN MODE (newtab) — V37: fixed overlay on col3 ---
      const col3Rect = window.Oracle.getCol3Rect();

      overlay = document.createElement('div');
      overlay.className = 'chat-slider-overlay';

      if (col3Rect) {
        overlay.style.cssText = `position: fixed; top: ${col3Rect.top}px; left: ${col3Rect.left}px; width: ${col3Rect.width}px; bottom: 0; z-index: 10000; display: flex; flex-direction: column; animation: fadeIn 0.15s ease-out;`;
      } else {
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);z-index:10000;display:flex;justify-content:flex-end;animation:fadeIn 0.2s ease-out;';
      }

      slider = document.createElement('div');
      slider.className = 'chat-slider';
      const sliderBorder = col3Rect ? `border-radius:12px;border:1px solid ${isDark ? 'rgba(255,255,255,0.08)' : 'rgba(225,232,237,0.6)'};` : '';
      slider.style.cssText = `width:100%;height:100%;background:${isDark ? '#1f2940' : 'white'};box-shadow:-4px 0 20px rgba(0,0,0,${isDark ? '0.4' : '0.15'});display:flex;flex-direction:column;animation:slideInRight 0.3s ease-out;${sliderBorder}`;

      slider.innerHTML = `
        <div style="padding:20px;border-bottom:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
          <div style="display:flex;align-items:center;gap:12px;">
            <div style="width:40px;height:40px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;">🤖</div>
            <div>
              <div style="font-weight:600;font-size:16px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Quick Chat</div>
              <div style="font-size:12px;color:${isDark ? '#888' : '#7f8c8d'};">Ask me anything</div>
            </div>
          </div>
          <button class="chat-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:18px;">×</button>
        </div>
        <div class="chat-messages" style="flex:1;overflow-y:auto;padding:20px;display:flex;flex-direction:column;gap:16px;">
          <div class="chat-welcome" style="text-align:center;padding:40px 20px;color:${isDark ? '#888' : '#7f8c8d'};"><div style="font-size:48px;margin-bottom:16px;">💬</div><div style="font-size:14px;">Send a message to start chatting</div></div>
        </div>
        <div style="padding:16px;border-top:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)'};flex-shrink:0;">
          <div style="display:flex;gap:8px;">
            <div class="chat-mention-container" style="position:relative;flex:1;display:flex;">
              <div class="chat-input" contenteditable="true" placeholder="Type your message... @name to tag people" style="flex:1;padding:12px 16px;background:${isDark ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.05)'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(102,126,234,0.2)'};border-radius:12px;font-size:14px;color:${isDark ? '#e8e8e8' : '#2c3e50'};max-height:120px;overflow-y:auto;outline:none;"></div>
            </div>
            <button class="chat-send-btn" style="background:linear-gradient(45deg,#667eea,#764ba2);border:none;color:white;width:44px;height:44px;border-radius:12px;cursor:pointer;font-size:18px;display:flex;align-items:center;justify-content:center;transition:all 0.2s;">➤</button>
          </div>
          <div style="font-size:11px;color:${isDark ? '#666' : '#95a5a6'};margin-top:8px;text-align:center;">Press Enter to send, Shift+Enter for new line · @name to tag people</div>
        </div>`;

      overlay.appendChild(slider);
      document.body.appendChild(overlay);

      chatInput = slider.querySelector('.chat-input');
      sendBtn = slider.querySelector('.chat-send-btn');
      messagesContainer = slider.querySelector('.chat-messages');

      setupMentionSupport(chatInput, slider.querySelector('.chat-mention-container'), isDark);

      const closeSlider = () => {
        overlay.style.animation = 'fadeOut 0.2s ease-out';
        slider.style.animation = 'slideOutRight 0.3s ease-out';
        document.removeEventListener('keydown', chatKeyHandler);
        state.isChatSliderOpen = false;
        if (onClose) onClose();
        setTimeout(() => { overlay.remove(); window.Oracle.collapseCol3AfterSlider(); }, 250);
      };
      slider.querySelector('.chat-close-btn').addEventListener('click', closeSlider);
      // V37: Click outside disabled — only Escape closes
      const chatKeyHandler = (e) => { if (e.key === 'Escape') closeSlider(); };
      document.addEventListener('keydown', chatKeyHandler);
    }

    // Placeholder behavior for contenteditable
    chatInput.addEventListener('focus', function () { if (this.textContent === '') this.dataset.placeholder = this.getAttribute('placeholder'); });

    let conversationHistory = [];

    // --- Unified send logic ---
    const sendMessage = async () => {
      const { api: message, display: displayMessage } = extractMessageWithMentions(chatInput);
      if (!message) return;
      messagesContainer.querySelector('.chat-welcome')?.remove();
      conversationHistory.push({ role: 'user', content: message, timestamp: new Date().toISOString() });

      const userMsgDiv = document.createElement('div');
      userMsgDiv.style.cssText = 'display:flex;justify-content:flex-end;';
      const userBubble = document.createElement('div');
      userBubble.style.cssText = `max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:${isInline ? '14px 14px 4px 14px' : '16px 16px 4px 16px'};font-size:${isInline ? '13px' : '14px'};line-height:1.5;`;
      // Show display text in bubble with mention names as styled chips
      // Split on @Name patterns that we know came from chips
      const escapedDisplay = escapeHtml(displayMessage);
      // Replace @Name mentions with styled chips — use a marker-based approach
      let bubbleHtml = escapedDisplay;
      // Find all @mentions in display text and style them
      bubbleHtml = bubbleHtml.replace(/@([A-Z][a-zA-Z]+(?:\s[A-Z][a-zA-Z]+)*)/g, 
        '<span style="background:rgba(255,255,255,0.2);padding:1px 6px;border-radius:4px;font-weight:600;">@$1</span>');
      userBubble.innerHTML = bubbleHtml;
      userMsgDiv.appendChild(userBubble);
      messagesContainer.appendChild(userMsgDiv);
      chatInput.innerHTML = '';

      const loadingDiv = document.createElement('div');
      loadingDiv.className = 'chat-loading';
      loadingDiv.style.cssText = 'display:flex;justify-content:flex-start;';
      loadingDiv.innerHTML = `<div style="padding:${isInline ? '10px 14px' : '12px 16px'};background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};border-radius:${isInline ? '14px 14px 14px 4px' : '16px 16px 16px 4px'};font-size:${isInline ? '13px' : '14px'};color:${isDark ? '#888' : '#7f8c8d'};"><span class="typing-dots">Thinking<span>.</span><span>.</span><span>.</span></span></div>`;
      messagesContainer.appendChild(loadingDiv);
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
      sendBtn.disabled = true; sendBtn.style.opacity = '0.6';

      // Bot message wrapper with action buttons
      const botMsgWrapper = document.createElement('div');
      botMsgWrapper.className = 'bot-msg-wrapper';
      botMsgWrapper.style.cssText = 'display:flex;flex-direction:column;align-items:flex-start;';
      const botMsgInner = document.createElement('div');
      botMsgInner.style.cssText = 'display:flex;justify-content:flex-start;width:100%;';
      const botBubble = document.createElement('div');
      botBubble.style.cssText = `max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};color:${isDark ? '#e8e8e8' : '#2c3e50'};border-radius:${isInline ? '14px 14px 14px 4px' : '16px 16px 16px 4px'};font-size:${isInline ? '13px' : '14px'};line-height:1.7;`;
      botMsgInner.appendChild(botBubble);
      botMsgWrapper.appendChild(botMsgInner);

      try {
        const responseText = await sendWithAblyStreaming({
          message, conversationHistory, messagesContainer, loadingDiv,
          botMsgDiv: botMsgWrapper, botBubble, isDark, source,
          onStreamDone: (finalText) => {
            const actions = createBotActions(finalText);
            botMsgWrapper.appendChild(actions);
          }
        });
        conversationHistory.push({ role: 'assistant', content: responseText, timestamp: new Date().toISOString() });
      } catch (error) {
        console.error('Chat stream error:', error);
        if (loadingDiv.parentNode) loadingDiv.remove();
        conversationHistory.push({ role: 'assistant', content: `Error: ${error.message}`, timestamp: new Date().toISOString(), error: true });
        const errDiv = document.createElement('div');
        errDiv.style.cssText = 'display:flex;justify-content:flex-start;';
        errDiv.innerHTML = `<div style="max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:rgba(231,76,60,0.1);color:#e74c3c;border-radius:${isInline ? '14px 14px 14px 4px' : '16px 16px 16px 4px'};font-size:${isInline ? '13px' : '14px'};line-height:1.5;">Error: ${error.message}</div>`;
        messagesContainer.appendChild(errDiv);
      }

      sendBtn.disabled = false; sendBtn.style.opacity = '1';
      messagesContainer.scrollTop = messagesContainer.scrollHeight;
    };

    sendBtn.addEventListener('click', sendMessage);
    chatInput.addEventListener('keydown', (e) => {
      const mentionDD = slider.querySelector('.chat-mention-dropdown.show');
      if (e.key === 'Enter' && !e.shiftKey && !mentionDD) { e.preventDefault(); sendMessage(); }
    });
    setTimeout(() => chatInput.focus(), 100);
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleAssistant = {
    showChatSlider,
    formatChatResponseWithAnnotations,
  };

})();
