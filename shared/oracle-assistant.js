// oracle-assistant.js — Quick Chat / Oracle assistant slider
// Supports two modes: 'fullscreen' (newtab) and 'inline' (sidepanel)
// V38: Native n8n webhook streaming (removed Ably dependency)
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
      .oracle-task-chip:hover{background:rgba(102,126,234,0.22)!important;border-color:rgba(102,126,234,0.4)!important;transform:translateY(-1px);}
      .oracle-task-chip:active{transform:translateY(0);background:rgba(102,126,234,0.3)!important;}
      body.dark-mode .oracle-task-chip{background:rgba(102,126,234,0.2)!important;color:#8fa4f0!important;border-color:rgba(102,126,234,0.3)!important;}
      body.dark-mode .oracle-task-chip:hover{background:rgba(102,126,234,0.35)!important;border-color:rgba(102,126,234,0.5)!important;}
      .oracle-source-chip:hover{background:rgba(102,126,234,0.22)!important;border-color:rgba(102,126,234,0.4)!important;transform:translateY(-1px);}
    `;
    document.head.appendChild(s);
  }

  // Ensure copy from anywhere in the extension copies as plain text
  document.addEventListener('copy', (e) => {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const container = sel.getRangeAt(0).commonAncestorContainer;
    const el = container.nodeType === 1 ? container : container.parentElement;
    // Force plain text copy from bot bubbles, chat sliders, transcript sliders, task cards, and Oracle UI
    const isOracleContent = el?.closest?.('.bot-msg-wrapper, .oracle-response, .chat-slider, .chat-slider-overlay, .slider-overlay, .transcript-slider-overlay, .conversation-messages, .transcript-messages, .container, .todos-container, .fyi-container');
    if (isOracleContent) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', sel.toString());
    }
  });

  // Ensure paste into any contenteditable input in the extension is plain text only
  document.addEventListener('paste', (e) => {
    const el = e.target;
    // Only intercept paste on contenteditable elements within Oracle UI (chat inputs, reply inputs)
    if (el?.isContentEditable || el?.closest?.('[contenteditable="true"]')) {
      const isOracleInput = el.closest?.('.chat-input, .chat-mention-container, .reply-input-container, .chat-slider, .slider-overlay, .transcript-slider-overlay, .container');
      if (isOracleInput) {
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData).getData('text/plain');
        document.execCommand('insertText', false, text);
      }
    }
  });

  // (Ably removed in V38 — streaming now handled via native n8n webhook streaming)

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
  function createBotActions(rawText, botBubbleEl) {
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

    function makeRichCopyBtn(title, svg) {
      const btn = document.createElement('button');
      btn.className = 'bot-action-btn';
      btn.title = title;
      btn.innerHTML = svg;
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        // Copy as rich HTML (for Google Docs table support) + plain text fallback
        const htmlContent = botBubbleEl ? botBubbleEl.innerHTML : '';
        const plainText = botBubbleEl ? botBubbleEl.innerText : rawText;
        try {
          const blob = new Blob([htmlContent], { type: 'text/html' });
          const textBlob = new Blob([plainText], { type: 'text/plain' });
          navigator.clipboard.write([
            new ClipboardItem({ 'text/html': blob, 'text/plain': textBlob })
          ]).then(() => {
            btn.classList.add('copied');
            btn.innerHTML = '✓';
            setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = svg; }, 1500);
          });
        } catch {
          // Fallback to plain text
          navigator.clipboard.writeText(plainText).then(() => {
            btn.classList.add('copied');
            btn.innerHTML = '✓';
            setTimeout(() => { btn.classList.remove('copied'); btn.innerHTML = svg; }, 1500);
          });
        }
      });
      return btn;
    }

    // Button 1: Copy links (only if links exist)
    if (links.length > 0) {
      actionsDiv.appendChild(makeCopyBtn('Copy links', linkSvg, links.join('\n')));
    }
    // Button 2: Copy full message as rich text (HTML tables paste properly into Google Docs)
    actionsDiv.appendChild(makeRichCopyBtn('Copy message', copySvg));

    return actionsDiv;
  }

  // ============================================
  // formatChatResponseWithAnnotations
  // ============================================
  // Task chip SVG icon (message/document icon)
  const taskChipSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="vertical-align:middle;flex-shrink:0;"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg>`;

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
      let iconHtml = linkIconSvg, title = 'Open link', isKnownService = false;
      const img = (src, alt, t) => { iconHtml = `<img src="${src}" alt="${alt}" style="width:14px;height:14px;vertical-align:middle;">`; title = t; isKnownService = true; };
      if (url.includes('slack.com')) img(slackIconUrl, 'Slack', 'Open in Slack');
      else if (url.includes('mail.google.com')) img(gmailIconUrl, 'Gmail', 'Open in Gmail');
      else if (url.includes('docs.google.com/document')) img(googleDocsIconUrl, 'Docs', 'Open in Google Docs');
      else if (url.includes('docs.google.com/spreadsheets') || url.includes('sheets.google.com')) img(googleSheetsIconUrl, 'Sheets', 'Open in Google Sheets');
      else if (url.includes('docs.google.com/presentation') || url.includes('slides.google.com')) img(googleSlidesIconUrl, 'Slides', 'Open in Google Slides');
      else if (url.includes('drive.google.com')) img(driveIconUrl, 'Drive', 'Open in Google Drive');
      else if (url.includes('freshdesk.com')) img(freshdeskIconUrl, 'Freshdesk', 'Open in Freshdesk');
      else if (url.includes('freshrelease.com')) img(freshreleaseIconUrl, 'Freshrelease', 'Open in Freshrelease');
      else if (url.includes('freshservice.com')) img(freshserviceIconUrl, 'Freshservice', 'Open in Freshservice');

      // For known services, show compact icon-only button
      if (isKnownService) {
        return `<a href="${url}" target="_blank" title="${title}" class="oracle-chat-link-icon" style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;background:rgba(102,126,234,0.1);border-radius:4px;margin:0 2px;text-decoration:none;vertical-align:middle;transition:all 0.2s;">${iconHtml}</a>`;
      }

      // For external URLs, show domain name + link icon as a chip
      let domain = '';
      try {
        domain = new URL(url).hostname.replace(/^www\./, '');
      } catch (e) {
        domain = url.replace(/^https?:\/\/(www\.)?/, '').split('/')[0];
      }
      return `<a href="${url}" target="_blank" title="${url}" class="oracle-source-chip" style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;background:rgba(102,126,234,0.1);color:#667eea;border-radius:10px;font-size:12px;font-weight:500;text-decoration:none;border:1px solid rgba(102,126,234,0.2);vertical-align:middle;transition:all 0.15s;margin:0 2px;">${linkIconSvg}<span style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(domain)}</span></a>`;
    }

    let result = text;

    // Normalize escaped newlines to real newlines (streaming may send \\n)
    result = result.replace(/\\n/g, '\n');

    // Convert [TASK:id] patterns to clickable task chips BEFORE other processing
    result = result.replace(/\[TASK:(\d+)\]/g, (match, taskId) => {
      return `%%TASKCHIP:${taskId}%%`;
    });

    // Convert [SOURCE:title|url] patterns to clickable source chips
    result = result.replace(/\[SOURCE:([^\]|]+)\|([^\]]+)\]/g, (match, title, url) => {
      return `%%SOURCECHIP:${encodeURIComponent(title)}|${encodeURIComponent(url)}%%`;
    });

    // Remove any incomplete [SOURCE: tags still being streamed (don't let URL regex corrupt them)
    result = result.replace(/\[SOURCE:[^\]]*$/g, '');

    // Remove inline "(Source: ... )" / "(Sources: ... )" / "(Source references: ... )" parenthetical text
    result = result.replace(/\(Source[s]?\s*(?:references?)?:[\s\S]*?\)\s*/gi, '');

    result = result.replace(/\[?(https?:\/\/[^\s\]\)]+)\]?/g, (match, url) => getIconForUrl(url));

    // Separate source chips that appear on their own line at the end → render as a References footer
    const sourceFooterChips = [];
    result = result.replace(/(?:\n|^)\s*(%%SOURCECHIP:[^%]+%%)\s*(?=\n|%%SOURCECHIP|$)/g, (match, chip) => {
      sourceFooterChips.push(chip);
      return '';
    });
    // Also collect any remaining inline source chips and move them to footer
    result = result.replace(/%%SOURCECHIP:[^%]+%%/g, (match) => {
      sourceFooterChips.push(match);
      return '';
    });

    // Clean up any remaining [SOURCE:...] tags that weren't fully parsed
    result = result.replace(/\[SOURCE:[^\]]*\]/g, '');
    // Clean up "References -" or "References:" lines from AI output (we render our own)
    result = result.replace(/\n\s*References?\s*[-:]?\s*\n/gi, '\n');

    // Convert markdown tables to HTML tables BEFORE escapeHtml
    const tablePlaceholders = [];
    result = result.replace(/(?:^|\n)((?:\|[^\n]+\|\s*\n){2,})/gm, (match, tableBlock) => {
      const lines = tableBlock.trim().split('\n').filter(l => l.trim());
      if (lines.length < 2) return match;
      // Check if second line is a separator (|---|---|)
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
      const isDark = document.body.classList.contains('dark-mode');
      const borderColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(102,126,234,0.15)';
      const headerBg = isDark ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.08)';
      const stripeBg = isDark ? 'rgba(255,255,255,0.03)' : 'rgba(102,126,234,0.03)';
      let html = `<div style="overflow-x:auto;margin:8px 0;"><table style="width:100%;border-collapse:collapse;font-size:13px;border:1px solid ${borderColor};border-radius:8px;overflow:hidden;">`;
      if (headerLine) {
        const cells = parseRow(headerLine);
        html += `<thead><tr style="background:${headerBg};">` + cells.map(c => `<th style="padding:8px 12px;text-align:left;font-weight:600;border-bottom:2px solid ${borderColor};white-space:nowrap;">${c}</th>`).join('') + '</tr></thead>';
      }
      html += '<tbody>';
      dataLines.forEach((line, i) => {
        const cells = parseRow(line);
        const bg = i % 2 === 1 ? stripeBg : 'transparent';
        html += `<tr style="background:${bg};">` + cells.map(c => `<td style="padding:6px 12px;border-bottom:1px solid ${borderColor};">${c}</td>`).join('') + '</tr>';
      });
      html += '</tbody></table></div>';
      const placeholder = `%%TABLE_${tablePlaceholders.length}%%`;
      tablePlaceholders.push(html);
      return '\n' + placeholder + '\n';
    });

    const parts = result.split(/(%%TASKCHIP:\d+%%|%%TABLE_\d+%%|<a href="[^"]*"[^>]*>.*?<\/a>)/g);
    result = parts.map(part => {
      if (part.startsWith('<a href=')) return part;
      if (part.startsWith('%%TABLE_')) {
        const idx = parseInt(part.match(/%%TABLE_(\d+)%%/)?.[1]);
        return tablePlaceholders[idx] || part;
      }
      if (part.startsWith('%%TASKCHIP:')) {
        const id = part.match(/%%TASKCHIP:(\d+)%%/)?.[1];
        if (id) return `<span class="oracle-task-chip" data-task-id="${id}" title="Open task #${id}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;background:rgba(102,126,234,0.12);color:#667eea;border-radius:6px;cursor:pointer;vertical-align:middle;transition:all 0.15s;border:1px solid rgba(102,126,234,0.2);margin:0 1px;">${taskChipSvg}</span>`;
        return part;
      }
      return escapeHtml(part);
    }).join('');
    // Render **bold** markdown
    result = result.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
    result = result.replace(/\n/g, '<br>');
    result = result.replace(/\s*-\s*(<a href)/g, ' $1');
    result = result.replace(/\s{2,}/g, ' ');

    // Append deduped References footer if there are sources
    if (sourceFooterChips.length > 0) {
      const seenUrls = new Set();
      const uniqueChips = sourceFooterChips.filter(c => {
        const m = c.match(/%%SOURCECHIP:[^|]+\|([^%]+)%%/);
        if (m) {
          const url = decodeURIComponent(m[1]);
          if (seenUrls.has(url)) return false;
          seenUrls.add(url);
          return true;
        }
        return true;
      });
      const redirectIconSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;flex-shrink:0;"><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><polyline points="15 3 21 3 21 9" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="10" y1="14" x2="21" y2="3" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
      const refItems = uniqueChips.map((c, i) => {
        const m = c.match(/%%SOURCECHIP:([^|]+)\|([^%]+)%%/);
        if (m) {
          const title = decodeURIComponent(m[1]);
          const url = decodeURIComponent(m[2]);
          return `<a href="${url}" target="_blank" title="${url}" style="display:flex;align-items:center;gap:6px;color:#667eea;text-decoration:none;font-size:12px;line-height:1.5;transition:opacity 0.15s;"><span style="color:#999;font-size:11px;min-width:14px;">${i + 1}.</span><span style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:280px;">${escapeHtml(title)}</span>${redirectIconSvg}</a>`;
        }
        return '';
      }).filter(Boolean).join('');
      result += `<div style="margin-top:12px;padding-top:10px;border-top:1px solid rgba(102,126,234,0.15);"><span style="font-size:11px;font-weight:600;color:#667eea;text-transform:uppercase;letter-spacing:0.5px;">References</span><div style="display:flex;flex-direction:column;gap:4px;margin-top:6px;">${refItems}</div></div>`;
    }

    return result;
  }

  function renderSourceChip(part) {
    const m = part.match(/%%SOURCECHIP:([^|]+)\|([^%]+)%%/);
    if (m) {
      const title = decodeURIComponent(m[1]);
      const url = decodeURIComponent(m[2]);
      const linkIconSvg = `<svg width="12" height="12" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="vertical-align:middle;flex-shrink:0;"><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71" stroke="#667eea" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
      return `<a href="${url}" target="_blank" title="${url}" class="oracle-source-chip" style="display:inline-flex;align-items:center;gap:4px;padding:2px 8px;background:rgba(102,126,234,0.1);color:#667eea;border-radius:10px;font-size:12px;font-weight:500;text-decoration:none;border:1px solid rgba(102,126,234,0.2);vertical-align:middle;transition:all 0.15s;margin:0 2px;">${linkIconSvg}<span style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${escapeHtml(title)}</span></a>`;
    }
    return part;
  }

  // ============================================
  // @mention search helper
  // ============================================
  let mentionAbortController = null;
  async function searchPeopleForMention(query) {
    if (mentionAbortController) mentionAbortController.abort();
    mentionAbortController = new AbortController();
    try {
      // For multi-word queries, search backend with first word for broader results
      const queryWords = query.toLowerCase().split(/\s+/).filter(w => w.length > 0);
      const backendQuery = queryWords.length > 1 ? queryWords[0] : query;
      const payload = {
        action: 'search_user_new_slack',
        query: backendQuery,
        platform: 'slack',
        context: 'mention',
        timestamp: new Date().toISOString(),
        source: 'oracle-assistant-mention',
        user_id: state.userData?.userId,
        authenticated: true
      };
      console.log('🔍 Mention search:', query, payload);
      const res = await fetch(WEBHOOK_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
        signal: mentionAbortController.signal
      });
      if (!res.ok) { console.warn('🔍 Mention search HTTP error:', res.status); return []; }
      let d = await res.json();
      console.log('🔍 Mention search raw response:', d);
      if (!Array.isArray(d)) d = d.results || d.members || d.users || [];
      // Filter to people (employees or items with email) — relaxed to allow missing type field
      const filtered = d.filter(r => {
        const t = (r.type || '').toLowerCase();
        const isPersonLike = !t || t === 'employee' || t === 'direct message' || t === 'dm' || t === 'user' || t === 'member';
        const hasIdentifier = r.user_email_ID || r.email || r.Full_Name || r['Full Name'] || r.full_name || r.name;
        return isPersonLike && hasIdentifier;
      });
      // Client-side multi-word filtering: every word in query must match name or email
      const multiWordFiltered = queryWords.length > 1
        ? filtered.filter(r => {
            const name = (r.Full_Name || r['Full Name'] || r.full_name || r.name || '').toLowerCase();
            const email = (r.user_email_ID || r.email || '').toLowerCase();
            const haystack = name + ' ' + email;
            return queryWords.every(w => haystack.includes(w));
          })
        : filtered;
      const finalResults = multiWordFiltered.slice(0, 5);
      console.log('🔍 Mention search filtered:', finalResults.length, 'results (query words:', queryWords, ')');
      return finalResults;
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

      // When caret is at element level (not inside a text node), determine which
      // child the offset points to so we know where to stop collecting text.
      const isElementLevel = range.startContainer.nodeType !== Node.TEXT_NODE;
      let stopNode = null;
      if (isElementLevel) {
        // startOffset = child index the caret sits before
        const children = range.startContainer.childNodes;
        stopNode = children[range.startOffset] || null; // null means caret is at the very end
      }

      const treeWalker = document.createTreeWalker(chatInput, NodeFilter.SHOW_TEXT, null);
      let fullText = '';
      let node;
      while (node = treeWalker.nextNode()) {
        if (!isElementLevel && node === range.startContainer) {
          // Caret is inside this text node — take text up to offset
          fullText += node.textContent.substring(0, range.startOffset);
          break;
        }
        // For element-level caret: stop when we reach or pass the stop node
        if (isElementLevel && stopNode) {
          // Check if this text node is inside or after the stopNode
          if (stopNode.contains(node) || node === stopNode) break;
          // Check if this text node comes after stopNode in document order
          if (stopNode.compareDocumentPosition(node) & Node.DOCUMENT_POSITION_FOLLOWING) break;
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
      // Walk backwards from end to find @ (allow spaces in names, like newtab.js reply handler)
      let atPos = -1;
      for (let i = textBefore.length - 1; i >= 0; i--) {
        if (textBefore[i] === '@') { atPos = i; break; }
        if (textBefore[i] === '\n') break;
      }

      if (atPos !== -1) {
        const query = textBefore.substring(atPos + 1);
        const trimmedQuery = query.trimEnd();
        mentionActive = true;
        console.log('🔍 @mention detected, query:', JSON.stringify(trimmedQuery), 'len:', trimmedQuery.length);
        if (trimmedQuery.length >= 2) {
          clearTimeout(mentionSearchTimeout);
          mentionSearchTimeout = setTimeout(async () => {
            const results = await searchPeopleForMention(trimmedQuery);
            renderMentionResults(results);
          }, 300);
        } else if (trimmedQuery.length === 0) {
          closeMentionDropdown();
        }
      } else {
        if (textBefore.includes('@')) console.log('🔍 @mention regex no match, textBefore ends:', JSON.stringify(textBefore.slice(-40)));
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
  // Bind click handlers on task chip elements
  // Opens transcript slider in newtab, or sends message to newtab from sidepanel
  // ============================================
  function bindTaskChips(container) {
    container.querySelectorAll('.oracle-task-chip:not([data-bound])').forEach(chip => {
      chip.dataset.bound = '1';
      chip.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        const taskId = chip.dataset.taskId;
        if (!taskId) return;

        // Newtab context: openTaskFromChat fetches task if needed, then opens slider
        if (typeof window.openTaskFromChat === 'function') {
          window.openTaskFromChat(taskId);
          return;
        }

        // Sidepanel context: try sidepanel's own showTranscriptSlider
        if (typeof showTranscriptSlider === 'function') {
          showTranscriptSlider(taskId);
          return;
        }

        // Fallback: send message to background to open in newtab
        try {
          chrome.runtime.sendMessage({ action: 'openTaskTranscript', taskId });
        } catch (err) {
          console.warn('Could not open task transcript:', err);
        }
      });
    });
  }

  // ============================================
  // Native n8n webhook streaming send logic
  // ============================================
  async function sendWithStreaming({ message, conversationHistory, messagesContainer, loadingDiv, botMsgDiv, botBubble, isDark, source, onStreamDone }) {
    const userId = await getFreshUserId();
    let fullResponseText = '', streamStarted = false;

    const controller = new AbortController();
    const streamTimeout = setTimeout(() => {
      controller.abort();
      if (!streamStarted) throw new Error('Response timeout');
    }, 120000);

    const response = await fetch(CHAT_WEBHOOK_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, conversation: conversationHistory, timestamp: new Date().toISOString(), source, user_id: userId }),
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
            // Only process streaming tokens from the AI Agent node, skip Respond to Webhook output
            if (parsed.type === 'item' && parsed.content && parsed.metadata?.nodeName !== 'Respond to Webhook') {
              if (!streamStarted) {
                streamStarted = true;
                if (loadingDiv.parentNode) loadingDiv.remove();
                messagesContainer.appendChild(botMsgDiv);
              }
              fullResponseText += parsed.content;
              try { botBubble.innerHTML = formatChatResponseWithAnnotations(fullResponseText); bindTaskChips(botBubble); } catch { botBubble.textContent = fullResponseText; }
              // No auto-scroll during streaming - user scrolls manually
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
      if (loadingDiv.parentNode) loadingDiv.remove();
      messagesContainer.appendChild(botMsgDiv);
    }

    try { botBubble.innerHTML = formatChatResponseWithAnnotations(fullResponseText); bindTaskChips(botBubble); } catch { botBubble.textContent = fullResponseText; }
    // Don't force-scroll at end - user may have been reading mid-response

    if (onStreamDone) onStreamDone(fullResponseText);
    return fullResponseText;
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
            <div style="width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;overflow:hidden;background:linear-gradient(45deg,#667eea,#764ba2);"><img src="${chrome.runtime.getURL('icon-oracle.png')}" style="width:32px;height:32px;border-radius:50%;object-fit:cover;"></div>
            <div>
              <div style="font-weight:600;font-size:14px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</div>
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
            <div style="width:36px;height:36px;border-radius:50%;display:flex;align-items:center;justify-content:center;overflow:hidden;background:linear-gradient(45deg,#667eea,#764ba2);"><img src="${chrome.runtime.getURL('icon-oracle.png')}" style="width:36px;height:36px;border-radius:50%;object-fit:cover;"></div>
            <div>
              <div style="font-weight:600;font-size:16px;color:${isDark ? '#e8e8e8' : '#2c3e50'};">Oracle Assistant</div>
              <div style="font-size:12px;color:${isDark ? '#888' : '#7f8c8d'};">Ask me anything</div>
            </div>
          </div>
          <div style="display:flex;align-items:center;gap:8px;">
            <button class="chat-expand-btn" title="Expand" style="background:${isDark ? 'rgba(102,126,234,0.2)' : 'rgba(102,126,234,0.1)'};border:none;color:#667eea;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:16px;display:flex;align-items:center;justify-content:center;transition:all 0.2s;">⬅</button>
            <button class="chat-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:18px;">×</button>
          </div>
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

      // Escape closes assistant slider first (before transcript behind it)
      const chatKeyHandler = (e) => {
        if (e.key === 'Escape' && state.isChatSliderOpen) {
          e.stopImmediatePropagation();
          e.preventDefault();
          closeSlider();
        }
      };
      document.addEventListener('keydown', chatKeyHandler, true); // capture phase

      const closeSlider = () => {
        document.removeEventListener('keydown', chatKeyHandler, true); // must match capture phase
        overlay.style.animation = 'fadeOut 0.2s ease-out';
        slider.style.animation = 'slideOutRight 0.3s ease-out';
        state.isChatSliderOpen = false;
        if (onClose) onClose();
        setTimeout(() => { overlay.remove(); window.Oracle.collapseCol3AfterSlider(); }, 250);
      };
      slider.querySelector('.chat-close-btn').addEventListener('click', closeSlider);

      // Expand/collapse toggle — 1.5x width like transcript
      const expandBtn = slider.querySelector('.chat-expand-btn');
      let isExpanded = false;
      if (expandBtn && col3Rect) {
        expandBtn.addEventListener('click', () => {
          isExpanded = !isExpanded;
          if (isExpanded) {
            const expandedWidth = Math.round(col3Rect.width * 1.5);
            overlay.style.width = expandedWidth + 'px';
            overlay.style.left = (col3Rect.left - (expandedWidth - col3Rect.width)) + 'px';
            expandBtn.innerHTML = '➡';
            expandBtn.title = 'Collapse';
          } else {
            overlay.style.width = col3Rect.width + 'px';
            overlay.style.left = col3Rect.left + 'px';
            expandBtn.innerHTML = '⬅';
            expandBtn.title = 'Expand';
          }
        });
      }
    }

    // Placeholder behavior for contenteditable
    chatInput.addEventListener('focus', function () { if (this.textContent === '') this.dataset.placeholder = this.getAttribute('placeholder'); });

    let conversationHistory = [];
    let _skipNextUserBubble = false;

    // --- Unified send logic ---
    const sendMessage = async () => {
      const { api: message, display: displayMessage } = extractMessageWithMentions(chatInput);
      if (!message) return;
      messagesContainer.querySelector('.chat-welcome')?.remove();
      conversationHistory.push({ role: 'user', content: message, timestamp: new Date().toISOString() });

      if (!_skipNextUserBubble) {
        const userMsgDiv = document.createElement('div');
        userMsgDiv.style.cssText = 'display:flex;justify-content:flex-end;';
        const userBubble = document.createElement('div');
        userBubble.style.cssText = `max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:${isInline ? '14px 14px 4px 14px' : '16px 16px 4px 16px'};font-size:${isInline ? '13px' : '14px'};line-height:1.5;`;
        const escapedDisplay = escapeHtml(displayMessage);
        let bubbleHtml = escapedDisplay;
        bubbleHtml = bubbleHtml.replace(/@([A-Z][a-zA-Z]+(?:\s[A-Z][a-zA-Z]+)*)/g, 
          '<span style="background:rgba(255,255,255,0.2);padding:1px 6px;border-radius:4px;font-weight:600;">@$1</span>');
        userBubble.innerHTML = bubbleHtml;
        userMsgDiv.appendChild(userBubble);
        messagesContainer.appendChild(userMsgDiv);
      }
      _skipNextUserBubble = false;
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
      botBubble.style.cssText = `max-width:${isInline ? '85%' : '95%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:${isDark ? 'rgba(102,126,234,0.15)' : 'rgba(102,126,234,0.08)'};color:${isDark ? '#e8e8e8' : '#2c3e50'};border-radius:${isInline ? '14px 14px 14px 4px' : '16px 16px 16px 4px'};font-size:${isInline ? '13px' : '14px'};line-height:1.7;`;
      botMsgInner.appendChild(botBubble);
      botMsgWrapper.appendChild(botMsgInner);

      try {
        const responseText = await sendWithStreaming({
          message, conversationHistory, messagesContainer, loadingDiv,
          botMsgDiv: botMsgWrapper, botBubble, isDark, source,
          onStreamDone: (finalText) => {
            const actions = createBotActions(finalText, botBubble);
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
      // No auto-scroll after response - user scrolls manually
    };

    sendBtn.addEventListener('click', sendMessage);
    chatInput.addEventListener('keydown', (e) => {
      const mentionDD = slider.querySelector('.chat-mention-dropdown.show');
      if (e.key === 'Enter' && !e.shiftKey && !mentionDD) { e.preventDefault(); sendMessage(); }
    });

    // Handle prefilled message from transcript Ably system-message — send directly
    if (opts.prefillMessage && !opts.showReadingState) {
      messagesContainer.querySelector('.chat-welcome')?.remove();
      // Show user bubble manually
      const userMsgDiv = document.createElement('div');
      userMsgDiv.style.cssText = 'display:flex;justify-content:flex-end;';
      const userBubble = document.createElement('div');
      userBubble.style.cssText = `max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:${isInline ? '14px 14px 4px 14px' : '16px 16px 4px 16px'};font-size:${isInline ? '13px' : '14px'};line-height:1.5;`;
      userBubble.textContent = opts.prefillMessage;
      userMsgDiv.appendChild(userBubble);
      messagesContainer.appendChild(userMsgDiv);
      // Set flag so sendMessage doesn't create a duplicate bubble
      _skipNextUserBubble = true;
      chatInput.textContent = opts.prefillMessage;
      setTimeout(() => sendMessage(), 100);
    }

    // Handle transcript "Reading the conversation..." state
    if (opts.showReadingState) {
      chatInput.setAttribute('contenteditable', 'false');
      chatInput.style.opacity = '0.6';
      sendBtn.disabled = true;
      sendBtn.style.opacity = '0.4';
      messagesContainer.querySelector('.chat-welcome')?.remove();
      const readingDiv = document.createElement('div');
      readingDiv.className = 'oracle-reading-state';
      readingDiv.style.cssText = 'display:flex;align-items:center;gap:10px;padding:16px;';
      readingDiv.innerHTML = `<div style="width:28px;height:28px;border:3px solid rgba(102,126,234,0.2);border-top-color:#667eea;border-radius:50%;animation:spin 0.8s linear infinite;flex-shrink:0;"></div><div style="color:${isDark ? '#888' : '#7f8c8d'};font-size:13px;">Reading the conversation...</div>`;
      messagesContainer.appendChild(readingDiv);

      // Store a resolver so we can update when extracted query arrives
      const sliderEl = isInline ? slider : overlay;
      sliderEl._oracleResolveQuery = (queryText) => {
        const readEl = messagesContainer.querySelector('.oracle-reading-state');
        if (readEl) readEl.remove();
        chatInput.setAttribute('contenteditable', 'true');
        chatInput.style.opacity = '1';
        sendBtn.disabled = false;
        sendBtn.style.opacity = '1';
        // Show user bubble manually
        const userMsgDiv = document.createElement('div');
        userMsgDiv.style.cssText = 'display:flex;justify-content:flex-end;';
        const userBubble = document.createElement('div');
        userBubble.style.cssText = `max-width:${isInline ? '85%' : '80%'};padding:${isInline ? '10px 14px' : '12px 16px'};background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:${isInline ? '14px 14px 4px 14px' : '16px 16px 4px 16px'};font-size:${isInline ? '13px' : '14px'};line-height:1.5;`;
        userBubble.textContent = queryText;
        userMsgDiv.appendChild(userBubble);
        messagesContainer.appendChild(userMsgDiv);
        // Set flag and send
        _skipNextUserBubble = true;
        chatInput.textContent = queryText;
        setTimeout(() => sendMessage(), 100);
      };
    }

    setTimeout(() => chatInput.focus(), 100);
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleAssistant = {
    showChatSlider,
    formatChatResponseWithAnnotations,
    bindTaskChips,
  };

})();
