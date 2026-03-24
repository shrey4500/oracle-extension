// oracle-message-format.js — Message formatting, sanitization, email iframe rendering, attachments
// Used by transcript slider and FYI display in both newtab and sidepanel

(function () {
  'use strict';

  const { escapeHtml, createAuthenticatedPayload } = window.Oracle;

  const GMAIL_ATTACHMENT_WEBHOOK = 'https://n8n-kqq5.onrender.com/webhook/gmail-attachment';

  // ============================================
  // EMOJI MAP (Slack codes → Unicode)
  // ============================================
  const emojiMap = {
    ':slightly_smiling_face:':'🙂',':smile:':'😄',':grinning:':'😀',':laughing:':'😆',':blush:':'😊',
    ':wink:':'😉',':heart_eyes:':'😍',':kissing_heart:':'😘',':thinking_face:':'🤔',':thinking:':'🤔',
    ':raised_hands:':'🙌',':clap:':'👏',':pray:':'🙏',':thumbsup:':'👍',':thumbsdown:':'👎',
    ':+1:':'👍',':-1:':'👎',':ok_hand:':'👌',':wave:':'👋',':point_up:':'☝️',
    ':point_down:':'👇',':point_left:':'👈',':point_right:':'👉',':fire:':'🔥',':star:':'⭐',
    ':sparkles:':'✨',':heart:':'❤️',':broken_heart:':'💔',':100:':'💯',':warning:':'⚠️',
    ':white_check_mark:':'✅',':x:':'❌',':heavy_check_mark:':'✔️',':question:':'❓',':exclamation:':'❗',
    ':eyes:':'👀',':see_no_evil:':'🙈',':hear_no_evil:':'🙉',':speak_no_evil:':'🙊',':rocket:':'🚀',
    ':tada:':'🎉',':party_popper:':'🎉',':bulb:':'💡',':memo:':'📝',':pencil:':'✏️',
    ':pushpin:':'📌',':calendar:':'📅',':clock:':'🕐',':hourglass:':'⏳',':email:':'📧',
    ':envelope:':'✉️',':phone:':'📞',':computer:':'💻',':link:':'🔗',':lock:':'🔒',
    ':key:':'🔑',':hammer:':'🔨',':wrench:':'🔧',':gear:':'⚙️',
    ':chart_with_upwards_trend:':'📈',':chart_with_downwards_trend:':'📉',':bar_chart:':'📊',
    ':moneybag:':'💰',':dollar:':'💵',':credit_card:':'💳',':gem:':'💎',':trophy:':'🏆',
    ':medal:':'🏅',':crown:':'👑',':muscle:':'💪',':brain:':'🧠',':robot_face:':'🤖',':robot:':'🤖',
    ':zap:':'⚡',':boom:':'💥',':collision:':'💥',':sweat_smile:':'😅',':joy:':'😂',
    ':sob:':'😭',':cry:':'😢',':angry:':'😠',':rage:':'😡',':sleepy:':'😪',':sleeping:':'😴',
    ':zzz:':'💤',':poop:':'💩',':ghost:':'👻',':skull:':'💀',':alien:':'👽',
    ':handshake:':'🤝',':fist:':'✊',':v:':'✌️',':crossed_fingers:':'🤞',':pinched_fingers:':'🤌',
    ':red_circle:':'🔴',':large_blue_circle:':'🔵',':green_circle:':'🟢',':yellow_circle:':'🟡',
    ':white_circle:':'⚪',':black_circle:':'⚫',':arrow_right:':'➡️',':arrow_left:':'⬅️',
    ':arrow_up:':'⬆️',':arrow_down:':'⬇️',':heavy_plus_sign:':'➕',':heavy_minus_sign:':'➖'
  };

  // ============================================
  // truncateUrl — Show clean, shortened URL for display
  // Extracts domain + path, resolves Google redirects, caps at ~60 chars
  // ============================================
  function truncateUrl(url) {
    if (!url || url.length <= 60) return url;
    try {
      const u = new URL(url);
      const domain = u.hostname.replace(/^www\./, '');
      const path = u.pathname;
      // If path is short enough, show domain + path
      if ((domain + path).length <= 55) {
        return domain + path + (u.search ? '…' : '');
      }
      // Otherwise show domain + truncated path
      const pathParts = path.split('/').filter(Boolean);
      if (pathParts.length <= 1) {
        return domain + '/' + (pathParts[0] || '').substring(0, 30) + '…';
      }
      // Show domain + first path segment + ... + last segment (truncated)
      const first = pathParts[0];
      const last = pathParts[pathParts.length - 1];
      const truncated = domain + '/' + first + '/…/' + (last.length > 20 ? last.substring(0, 20) + '…' : last);
      return truncated.length > 65 ? domain + '/…/' + (last.length > 25 ? last.substring(0, 25) + '…' : last) : truncated;
    } catch {
      // Fallback: simple character truncation
      return url.substring(0, 55) + '…';
    }
  }

  // ============================================
  // formatMessageContent — Slack/plain text formatting
  // ============================================
  function formatMessageContent(text) {
    if (!text) return '';

    // Pre-process Slack link patterns BEFORE HTML detection
    // These look like HTML tags (<url|text>, <mailto:...>) but are Slack formatting
    let processed = text;

    // Slack user mentions <@U1234|Name> or <@U1234> → styled chip placeholder
    processed = processed.replace(/<@([A-Z0-9]+)\|([^>]+)>/g, '%%MENTION%%$2%%ENDMENTION%%');
    processed = processed.replace(/<@([A-Z0-9]+)>/g, '%%MENTION%%$1%%ENDMENTION%%');
    // Slack channel mentions <!subteam^ID|@Name> or <!here> etc
    processed = processed.replace(/<!subteam\^[A-Z0-9]+\|@([^>]+)>/g, '%%MENTION%%$1%%ENDMENTION%%');
    processed = processed.replace(/<!([a-z]+)>/g, '%%MENTION%%$1%%ENDMENTION%%');

    // Slack <url|text> → markdown-style placeholder to preserve through HTML detection
    processed = processed.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, '%%SLACKLINK:$1%%$2%%ENDSLACKLINK%%');
    // Slack <mailto:email|text>
    processed = processed.replace(/<mailto:([^|>]+)\|([^>]+)>/g, '%%SLACKMAILTO:$1%%$2%%ENDSLACKMAILTO%%');
    // Slack <mailto:email>
    processed = processed.replace(/<mailto:([^>]+)>/g, '%%SLACKMAILTO:$1%%$1%%ENDSLACKMAILTO%%');
    // Slack <url> (bare)
    processed = processed.replace(/<(https?:\/\/[^>]+)>/g, '%%SLACKLINK:$1%%$1%%ENDSLACKLINK%%');

    // HTML content → sanitize (now safe since Slack links are placeholdered)
    if (/<[a-z][\s\S]*>/i.test(processed)) {
      // Restore Slack links before sanitizing
      processed = processed.replace(/%%SLACKLINK:(.*?)%%(.*?)%%ENDSLACKLINK%%/g, (_, url, label) => {
        const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
        return `<a href="${url}" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">${displayText}</a>`;
      });
      processed = processed.replace(/%%SLACKMAILTO:(.*?)%%(.*?)%%ENDSLACKMAILTO%%/g, '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$2</a>');
      // Also convert Markdown links in HTML content
      processed = processed.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, (_, label, url) => {
        const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
        return `<a href="${url}" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">${displayText}</a>`;
      });
      processed = processed.replace(/\[([^\]]+)\]\(mailto:([^)]+)\)/g, '<a href="mailto:$2" style="color:#667eea;text-decoration:underline;">$1</a>');
      // Restore @mention chips in HTML content
      processed = processed.replace(/%%MENTION%%(.*?)%%ENDMENTION%%/g, '<span style="display:inline-flex;align-items:center;gap:2px;background:linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1));color:#667eea;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;">@$1</span>');
      return sanitizeHtml(processed);
    }

    // Restore Slack links for plain text path (they'll be handled below)
    processed = processed.replace(/%%SLACKLINK:(.*?)%%(.*?)%%ENDSLACKLINK%%/g, '<$1|$2>');
    processed = processed.replace(/%%SLACKMAILTO:(.*?)%%(.*?)%%ENDSLACKMAILTO%%/g, '<mailto:$1|$2>');

    // Plain text
    let fmt = processed;

    // Decode HTML entities
    fmt = fmt.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'");

    // Convert Slack link patterns to placeholders BEFORE escaping HTML
    // Slack <url|text>
    fmt = fmt.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, (_, url, label) => {
      const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
      return `%%LINK_A%%${url}%%LINK_B%%${escapeHtml(displayText)}%%LINK_T%%${escapeHtml(label)}%%LINK_END%%`;
    });
    // Slack <mailto:email|text>
    fmt = fmt.replace(/<mailto:([^|>]+)\|([^>]+)>/g, '%%MAILTO_A%%$1%%MAILTO_B%%$2%%MAILTO_END%%');
    // Slack <mailto:email>
    fmt = fmt.replace(/<mailto:([^>]+)>/g, '%%MAILTO_A%%$1%%MAILTO_B%%$1%%MAILTO_END%%');
    // Slack <url> bare
    fmt = fmt.replace(/<(https?:\/\/[^>]+)>/g, (_, url) => {
      return `%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });

    // Markdown links [text](url) — handle BEFORE escapeHtml to avoid & encoding issues
    fmt = fmt.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, (_, label, url) => {
      const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : escapeHtml(label);
      return `%%LINK_A%%${url}%%LINK_B%%${displayText}%%LINK_T%%${escapeHtml(label)}%%LINK_END%%`;
    });
    // Markdown mailto links [text](mailto:email)
    fmt = fmt.replace(/\[([^\]]+)\]\(mailto:([^)]+)\)/g, '%%MAILTO_A%%$2%%MAILTO_B%%$1%%MAILTO_END%%');

    // Plain URLs — convert BEFORE escapeHtml to avoid & encoding issues
    fmt = fmt.replace(/(^|[\s\n])(https?:\/\/[^\s<>)\]]+)/g, (_, prefix, url) => {
      return `${prefix}%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });

    // Escape HTML (safe now that all links are placeholdered)
    fmt = escapeHtml(fmt);

    // Restore Slack link placeholders as proper <a> tags
    fmt = fmt.replace(/%%LINK_A%%(.*?)%%LINK_B%%(.*?)%%LINK_T%%(.*?)%%LINK_END%%/g,
      '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;" title="$3">$2</a>');
    fmt = fmt.replace(/%%MAILTO_A%%(.*?)%%MAILTO_B%%(.*?)%%MAILTO_END%%/g,
      '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$2</a>');

    // Restore @mention chips (from Slack <@U123|Name> pre-processing)
    fmt = fmt.replace(/%%MENTION%%(.*?)%%ENDMENTION%%/g, '<span style="display:inline-flex;align-items:center;gap:2px;background:linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1));color:#667eea;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;">@$1</span>');
    // Also style plain @Name mentions (e.g., @Shrey Jain) — match @FirstName LastName pattern
    // Match 1-3 capitalized words separated by single spaces (not newlines).
    // Use a greedy match then trim trailing common English words that aren't part of names.
    fmt = fmt.replace(/@([A-Z][a-z]+(?:[ ][A-Z][a-z]+){0,2})/g, (match, name) => {
      // Trim trailing words that are likely sentence starters, not name parts
      const commonWords = /\s+(Could|Would|Should|Please|Can|Will|May|Might|The|This|That|What|When|Where|How|Why|Who|Hi|Hello|Hey|Thanks|Thank|Also|But|And|Or|For|From|With|Just|Do|Does|Did|Has|Have|Had|Is|Are|Was|Were|Be|Not|So|If|As|At|In|On|To|Let|See|Get|Got|Need|Help|Share|Check|Sure|Your|Our|My|Its|His|Her)$/;
      const cleaned = name.replace(commonWords, '');
      return '<span style="display:inline-flex;align-items:center;gap:2px;background:linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1));color:#667eea;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;">@' + cleaned + '</span>' + name.slice(cleaned.length);
    });

    // Slack emoji codes
    Object.keys(emojiMap).forEach(code => { fmt = fmt.split(code).join(emojiMap[code]); });
    fmt = fmt.replace(/:([a-z0-9_+-]+):/g, (m) => emojiMap[m] || m);

    // Code blocks (multiline: ```...``` with dotAll flag)
    fmt = fmt.replace(/```([\s\S]+?)```/g,
      '<pre style="background:rgba(45,55,72,0.08);border:1px solid rgba(45,55,72,0.15);border-radius:6px;padding:12px;margin:8px 0;font-family:\'SF Mono\',Monaco,\'Courier New\',monospace;font-size:12px;overflow-x:hidden;white-space:pre-wrap;word-break:break-all;max-width:100%;">$1</pre>');
    // Inline code
    fmt = fmt.replace(/`([^`\n]+)`/g,
      '<code style="background:rgba(45,55,72,0.08);border-radius:4px;padding:2px 6px;font-family:\'SF Mono\',Monaco,\'Courier New\',monospace;font-size:12px;color:#e53e3e;">$1</code>');

    // Emails (still needs to run post-escape since it's simple pattern matching)
    fmt = fmt.replace(/(?<!href="mailto:|">)(?<![a-zA-Z0-9._%+-])([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})(?!<\/a>)/g, '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$1</a>');

    // Bold *text*
    fmt = fmt.replace(/\*([^*\n]+)\*/g, '<strong style="font-weight:600;color:#2c3e50;">$1</strong>');
    // Italic _text_
    fmt = fmt.replace(/(?<!\w)_([^_\n]+)_(?!\w)/g, '<em style="font-style:italic;">$1</em>');
    // Strikethrough ~text~
    fmt = fmt.replace(/~([^~\n]+)~/g, '<del style="text-decoration:line-through;opacity:0.7;">$1</del>');

    // Numbered headers
    fmt = fmt.replace(/^(\d+\.\s*[^:\n]+:)/gm, '<strong style="font-weight:600;color:#2c3e50;display:block;margin-top:12px;margin-bottom:4px;">$1</strong>');
    // Numbered list items
    fmt = fmt.replace(/^(\d+)\.\s+/gm, '<span style="color:#667eea;font-weight:600;margin-right:4px;">$1.</span> ');
    // Remove [image: ...] placeholders
    fmt = fmt.replace(/\[image:[^\]]+\]/g, '');

    // Blockquotes (> prefix)
    const lines = fmt.split('<br>');
    let inQuote = false, quoteLines = [], result = [];
    for (const line of lines) {
      if (/^(&gt;|>)\s?/.test(line.trim())) {
        quoteLines.push(line.trim().replace(/^(&gt;|>)\s?/, ''));
        inQuote = true;
      } else {
        if (inQuote && quoteLines.length > 0) {
          result.push(`<blockquote style="border-left:4px solid #667eea;margin:8px 0;padding:8px 12px;background:rgba(102,126,234,0.08);border-radius:0 8px 8px 0;color:inherit;font-style:normal;">${quoteLines.join('<br>')}</blockquote>`);
          quoteLines = []; inQuote = false;
        }
        result.push(line);
      }
    }
    if (quoteLines.length > 0) {
      result.push(`<blockquote style="border-left:4px solid #667eea;margin:8px 0;padding:8px 12px;background:rgba(102,126,234,0.08);border-radius:0 8px 8px 0;color:inherit;font-style:normal;">${quoteLines.join('<br>')}</blockquote>`);
    }
    fmt = result.join('<br>');

    // Cleanup
    fmt = fmt.replace(/<br>{3,}/g, '<br><br>');
    // Newlines outside <pre> blocks
    const preParts = fmt.split(/(<pre[^>]*>[\s\S]*?<\/pre>)/);
    fmt = preParts.map(p => p.startsWith('<pre') ? p : p.replace(/\n/g, '<br>')).join('');

    return fmt;
  }

  // ============================================
  // sanitizeHtml — Strip dangerous elements, style safe elements
  // ============================================
  function sanitizeHtml(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;

    temp.querySelectorAll('script, style, iframe, object, embed').forEach(el => el.remove());

    temp.querySelectorAll('*').forEach(el => {
      Array.from(el.attributes).forEach(attr => {
        if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) {
          el.removeAttribute(attr.name);
        }
      });
    });

    temp.querySelectorAll('table').forEach(t => { t.style.cssText = 'border-collapse:collapse;width:100%;margin:10px 0;font-size:13px;overflow-x:hidden;display:block;max-width:100%;table-layout:fixed;'; });
    temp.querySelectorAll('th').forEach(th => { th.style.cssText = 'background:linear-gradient(45deg,#667eea,#764ba2);color:white;padding:10px 12px;text-align:left;font-weight:600;border:1px solid rgba(102,126,234,0.3);white-space:nowrap;'; });
    temp.querySelectorAll('td').forEach(td => { td.style.cssText = 'padding:8px 12px;border:1px solid rgba(225,232,237,0.6);vertical-align:top;'; });
    temp.querySelectorAll('tr').forEach((tr, i) => { if (i % 2 === 1) tr.style.backgroundColor = 'rgba(102,126,234,0.03)'; });
    temp.querySelectorAll('a').forEach(a => { a.style.cssText = 'color:#667eea;text-decoration:underline;'; a.setAttribute('target', '_blank'); });
    temp.querySelectorAll('p').forEach(p => { p.style.cssText = 'margin:8px 0;line-height:1.5;'; });
    temp.querySelectorAll('b, strong').forEach(b => { b.style.cssText = 'font-weight:600;color:#2c3e50;'; });

    return temp.innerHTML;
  }

  // ============================================
  // isComplexEmailHtml — Detect complex email (nested tables, signatures)
  // ============================================
  function isComplexEmailHtml(html) {
    if (!html) return false;
    const temp = document.createElement('div');
    temp.innerHTML = html;
    const nestedTables = temp.querySelectorAll('table table');
    const sigs = temp.querySelectorAll('[id*="Signature"], [id*="signature"], .elementToProof');
    const hrs = temp.querySelectorAll('hr');
    const hasMany = (html.match(/style="/g) || []).length > 10;
    return nestedTables.length > 0 || sigs.length > 0 || (hrs.length > 0 && hasMany);
  }

  // ============================================
  // stripEmailQuotedContent — Remove quoted email chain content
  // ============================================
  function stripEmailQuotedContent(html) {
    if (!html) return html;
    const temp = document.createElement('div');
    temp.innerHTML = html;

    // Remove Gmail quoted content containers
    temp.querySelectorAll('.gmail_quote, .gmail_quote_container').forEach(el => el.remove());

    // Remove Outlook-style quoted content (reference messages)
    temp.querySelectorAll('#m_[id*="mail-editor-reference-message-container"], [id*="reference-message-container"]').forEach(el => el.remove());
    temp.querySelectorAll('div[id*="mail-editor-reference-message"]').forEach(el => el.remove());

    // Remove email signatures
    temp.querySelectorAll('.gmail_signature, .gmail_signature_prefix').forEach(el => el.remove());

    // Remove confidentiality disclaimers (common in enterprise emails)
    temp.querySelectorAll('b > i > span').forEach(el => {
      if (el.textContent && el.textContent.includes('confidential') && el.textContent.length > 100) {
        // Remove the entire bold/italic wrapper
        const parent = el.closest('b');
        if (parent) parent.remove();
      }
    });

    // Remove trailing <br> and empty divs
    const result = temp.innerHTML
      .replace(/(<br\s*\/?>[\s\n]*)+$/gi, '')
      .replace(/(<div>\s*<\/div>\s*)+$/gi, '')
      .trim();

    return result || html; // Fallback to original if stripping removed everything
  }

  // ============================================
  // renderEmailInIframe — Sandboxed iframe for email HTML
  // ============================================
  function renderEmailInIframe(html, container, isDark) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    temp.querySelectorAll('script, object, embed').forEach(el => el.remove());
    temp.querySelectorAll('*').forEach(el => {
      Array.from(el.attributes).forEach(attr => {
        if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) el.removeAttribute(attr.name);
      });
      // Only strip min-width constraints, keep actual widths for proper table layout
      const style = el.getAttribute('style') || '';
      if (style) {
        el.setAttribute('style', style.replace(/min-width\s*:\s*\d{3,}px/gi, 'min-width:0'));
      }
    });
    temp.querySelectorAll('a').forEach(link => { link.setAttribute('target', '_blank'); link.setAttribute('rel', 'noopener noreferrer'); });
    const cleanHtml = temp.innerHTML;

    const textColor = isDark ? '#e8e8e8' : '#2c3e50';
    const linkColor = isDark ? '#8b9ff0' : '#667eea';

    const iframeDoc = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box}
html{overflow:hidden}
body{margin:0;padding:8px 12px;background:transparent;color:${textColor};font-family:'Lato',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:14px;line-height:1.6;word-wrap:break-word;overflow-wrap:break-word;overflow:hidden;height:auto}
a{color:${linkColor};word-break:break-all}
img{max-width:100%!important;height:auto!important}
table{border-collapse:collapse;overflow:visible}
td,th{word-break:normal;overflow-wrap:break-word}
div,span,p,section,article{overflow-wrap:break-word}
pre{white-space:pre-wrap;word-break:break-all;overflow-x:auto;max-width:100%}
blockquote{margin:8px 0;padding:8px 12px;border-left:3px solid ${linkColor};background:${isDark?'rgba(102,126,234,0.1)':'rgba(102,126,234,0.05)'};max-width:100%;overflow:auto}
hr{border:none;border-top:1px solid ${isDark?'rgba(255,255,255,0.12)':'rgba(0,0,0,0.08)'};margin:10px 0}
table[cellpadding],table[cellspacing]{font-size:13px}
h2{font-size:15px;margin:4px 0}p{margin:4px 0}
ul,ol{margin:4px 0;padding-left:20px}li{margin:2px 0}
b,strong{font-weight:600;color:${textColor}}
img[alt="mobilePhone"],img[alt="emailAddress"],img[alt="website"],img[alt="address"]{width:12px!important;height:auto}
div[class*="elementToProof"]{margin-top:0.4em;margin-bottom:0.4em}
</style></head><body>${cleanHtml}</body></html>`;

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'width:100%;border:none;min-height:60px;background:transparent;display:block;margin:0;padding:0;';
    iframe.setAttribute('sandbox', 'allow-same-origin allow-popups allow-popups-to-escape-sandbox');
    iframe.srcdoc = iframeDoc;

    iframe.onload = () => {
      try {
        const resize = () => {
          if (iframe.contentDocument?.body) {
            // Use the max of body.scrollHeight and documentElement.scrollHeight for accuracy
            const bodyH = iframe.contentDocument.body.scrollHeight;
            const docH = iframe.contentDocument.documentElement.scrollHeight;
            const newHeight = Math.max(bodyH, docH) + 8;
            // Only update if height actually changed and is reasonable
            if (newHeight > 60) {
              iframe.style.height = newHeight + 'px';
            }
          }
        };
        resize();
        iframe.contentDocument.querySelectorAll('img').forEach(img => { if (!img.complete) img.addEventListener('load', resize); });
        setTimeout(resize, 100);
        setTimeout(resize, 300);
        setTimeout(resize, 600);
        setTimeout(resize, 1000);
        setTimeout(resize, 2000);
        // Use ResizeObserver on the iframe body for dynamic content
        if (typeof ResizeObserver !== 'undefined' && iframe.contentDocument?.body) {
          const ro = new ResizeObserver(resize);
          ro.observe(iframe.contentDocument.body);
        }
      } catch (e) { /* cross-origin iframe safety */ }
    };

    container.style.padding = '0';
    container.innerHTML = '';
    container.appendChild(iframe);
  }

  // ============================================
  // fetchGmailAttachment — Fetch via webhook
  // ============================================
  async function fetchGmailAttachment(messageId, attachmentId, mimeType, filename) {
    try {
      if (typeof showToastNotification === 'function') showToastNotification('Fetching image...');
      const response = await fetch(GMAIL_ATTACHMENT_WEBHOOK, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'fetch_gmail_attachment',
          messageId, attachmentId, mimeType, filename,
          timestamp: new Date().toISOString()
        }))
      });
      const data = await response.json();
      if (data.success && data.base64Data) {
        showAttachmentPreview(`data:${mimeType};base64,${data.base64Data}`, mimeType, filename);
      } else {
        if (typeof showToastNotification === 'function') showToastNotification(data.error || 'Failed to fetch attachment');
      }
    } catch (e) {
      console.error('Error fetching Gmail attachment:', e);
      if (typeof showToastNotification === 'function') showToastNotification('Failed to fetch attachment');
    }
  }

  // ============================================
  // renderTranscriptAttachment — Build compact attachment element
  // ============================================
  function renderTranscriptAttachment(attachment) {
    const { name, type, url, size, text, attachmentId, messageId, isInline, contentId } = attachment;
    if (!name && !url && !text && !attachmentId) return null;

    const isDark = document.body.classList.contains('dark-mode');
    const attachEl = document.createElement('div');
    const isGmailAttachment = attachmentId && messageId;
    const lt = (type || '').toLowerCase();
    const ln = (name || '').toLowerCase();
    const lu = (url || '').toLowerCase();

    // Determine icon
    let icon = '📎', iconBg = 'linear-gradient(45deg,#667eea,#764ba2)';
    const isImage = lt.includes('image') || lt === 'jpeg' || lt === 'png' || lt === 'gif' ||
      ln.endsWith('.png') || ln.endsWith('.jpg') || ln.endsWith('.jpeg') || ln.endsWith('.gif') || ln.endsWith('.webp');
    const isPreviewable = isImage || lt.includes('video') || lt.includes('movie') || lt.includes('mpeg') ||
      lt.includes('pdf') || ln.endsWith('.mp4') || ln.endsWith('.mov') || ln.endsWith('.webm') || ln.endsWith('.pdf') ||
      lu.includes('.png') || lu.includes('.jpg') || lu.includes('.jpeg') || lu.includes('.mp4') || lu.includes('.pdf');

    const isLinkPreview = lt === 'link' || lt.includes('linkedin') || lt.includes('youtube') || lt.includes('google docs') ||
      lt.includes('google sheets') || lt.includes('google slides') || lt.includes('attio') || lt.includes('openai') ||
      lt.includes('freshworks') || (lt.includes('.com') || lt.includes('.in') || lt.includes('.io'));

    if (lt.includes('video') || lt.includes('movie') || lt.includes('mpeg') || ln.endsWith('.mov') || ln.endsWith('.mp4') || ln.endsWith('.avi') || ln.endsWith('.webm')) { icon = '🎬'; iconBg = 'linear-gradient(45deg,#e74c3c,#c0392b)'; }
    else if (lt.includes('youtube')) { icon = '▶️'; iconBg = 'linear-gradient(45deg,#ff0000,#cc0000)'; }
    else if (lt.includes('linkedin')) { icon = '💼'; iconBg = 'linear-gradient(45deg,#0077b5,#005885)'; }
    else if (lt.includes('audio') || ln.endsWith('.mp3') || ln.endsWith('.wav') || ln.endsWith('.m4a')) { icon = '🎵'; iconBg = 'linear-gradient(45deg,#9b59b6,#8e44ad)'; }
    else if (isImage) { icon = '🖼️'; iconBg = 'linear-gradient(45deg,#3498db,#2980b9)'; }
    else if (lt.includes('pdf') || ln.endsWith('.pdf')) { icon = '📄'; iconBg = 'linear-gradient(45deg,#e74c3c,#c0392b)'; }
    else if (lt.includes('google docs') || lt.includes('word') || lt.includes('document') || ln.endsWith('.doc') || ln.endsWith('.docx')) { icon = '📝'; iconBg = 'linear-gradient(45deg,#4285f4,#2a5db0)'; }
    else if (lt.includes('google sheets') || lt.includes('sheet') || lt.includes('excel') || ln.endsWith('.xls') || ln.endsWith('.xlsx') || ln.endsWith('.csv')) { icon = '📊'; iconBg = 'linear-gradient(45deg,#0f9d58,#0b7a45)'; }
    else if (lt.includes('google slides') || lt.includes('presentation') || lt.includes('powerpoint') || ln.endsWith('.ppt') || ln.endsWith('.pptx')) { icon = '📽️'; iconBg = 'linear-gradient(45deg,#f4b400,#c99200)'; }
    else if (lt.includes('zip') || lt.includes('archive') || ln.endsWith('.zip') || ln.endsWith('.rar') || ln.endsWith('.7z')) { icon = '📦'; iconBg = 'linear-gradient(45deg,#7f8c8d,#5d6d7e)'; }
    else if (isLinkPreview) { icon = '🔗'; iconBg = 'linear-gradient(45deg,#667eea,#764ba2)'; }

    let sizeStr = '';
    if (size) {
      if (size < 1024) sizeStr = size + ' B';
      else if (size < 1024 * 1024) sizeStr = (size / 1024).toFixed(1) + ' KB';
      else sizeStr = (size / (1024 * 1024)).toFixed(1) + ' MB';
    }

    let previewText = text && text.length > 0 ? (text.length > 150 ? text.substring(0, 150) + '...' : text) : '';
    const hasPreview = text && text.length > 0 && isLinkPreview;

    // Link previews stay full-width
    if (hasPreview) {
      attachEl.style.cssText = `display:flex;flex-direction:column;gap:8px;padding:12px;background:${isDark ? 'rgba(255,255,255,0.04)' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'};border-radius:10px;transition:all 0.2s;cursor:pointer;width:100%;`;
      attachEl.innerHTML = `
        <div style="display:flex;align-items:center;gap:12px;">
          <div style="width:36px;height:36px;background:${iconBg};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;">${icon}</div>
          <div style="flex:1;min-width:0;overflow:hidden;">
            <div style="font-weight:600;font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Link')}">${escapeHtml(name || 'Link')}</div>
            <div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">${type ? escapeHtml(type) : 'Link'}</div>
          </div>
          ${url ? `<a href="${escapeHtml(url)}" target="_blank" style="display:flex;align-items:center;justify-content:center;width:28px;height:28px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:6px;color:white;text-decoration:none;font-size:12px;flex-shrink:0;" title="Open link">↗</a>` : ''}
        </div>
        <div style="font-size:12px;color:${isDark ? '#aaa' : '#5d6d7e'};line-height:1.5;padding-left:48px;border-left:3px solid rgba(102,126,234,0.3);margin-left:16px;">${escapeHtml(previewText)}</div>`;
    } else {
      // Compact card layout — fits multiple per row
      const truncName = (name || 'Attachment').length > 18 ? (name || 'Attachment').substring(0, 15) + '...' : (name || 'Attachment');

      // For images with URL, show thumbnail; for Gmail attachments, show icon with fetch-on-click
      let thumbnailHtml;
      if (isImage && url) {
        thumbnailHtml = `<div class="attach-thumb" style="width:48px;height:48px;border-radius:8px;overflow:hidden;flex-shrink:0;background:${isDark ? 'rgba(255,255,255,0.05)' : '#f0f2f5'};"><img src="${escapeHtml(url)}" style="width:100%;height:100%;object-fit:cover;" data-fallback-icon="${icon}" data-fallback-bg="${iconBg}"></div>`;
      } else if (isImage && isGmailAttachment) {
        thumbnailHtml = `<div class="attach-thumb gmail-fetch-btn" data-message-id="${escapeHtml(messageId)}" data-attachment-id="${escapeHtml(attachmentId)}" data-mime-type="${escapeHtml(type || 'image/png')}" data-filename="${escapeHtml(name || 'attachment')}" style="width:48px;height:48px;border-radius:8px;overflow:hidden;flex-shrink:0;background:${iconBg};display:flex;align-items:center;justify-content:center;font-size:22px;cursor:pointer;" title="Click to load">${icon}</div>`;
      } else {
        thumbnailHtml = `<div style="width:40px;height:40px;background:${iconBg};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;">${icon}</div>`;
      }

      attachEl.style.cssText = `display:flex;align-items:center;gap:10px;padding:8px 10px;background:${isDark ? 'rgba(255,255,255,0.04)' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'};border-radius:10px;transition:all 0.2s;cursor:pointer;min-width:0;flex:1 1 auto;max-width:260px;`;

      attachEl.innerHTML = `
        ${thumbnailHtml}
        <div style="flex:1;min-width:0;overflow:hidden;">
          <div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Attachment')}">${escapeHtml(truncName)}</div>
          <div style="font-size:10px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:1px;">${sizeStr || (type ? escapeHtml(type) : '')}</div>
        </div>`;
    }

    // Image thumbnail fallback (CSP-safe, no inline onerror)
    const thumbImg = attachEl.querySelector('.attach-thumb img[data-fallback-icon]');
    if (thumbImg) {
      thumbImg.addEventListener('error', function() {
        const icon = this.dataset.fallbackIcon;
        const bg = this.dataset.fallbackBg;
        this.parentElement.innerHTML = `<div style="width:100%;height:100%;background:${bg};display:flex;align-items:center;justify-content:center;font-size:22px;">${icon}</div>`;
      });
    }

    // Hover effects
    attachEl.addEventListener('mouseenter', () => { attachEl.style.background = isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'; attachEl.style.borderColor = 'rgba(102,126,234,0.3)'; });
    attachEl.addEventListener('mouseleave', () => { attachEl.style.background = isDark ? 'rgba(255,255,255,0.04)' : 'white'; attachEl.style.borderColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'; });

    // Click handlers
    if (isGmailAttachment) {
      const gmailBtn = attachEl.querySelector('.gmail-fetch-btn');
      const fetchGmail = (el) => {
        fetchGmailAttachment(el.getAttribute('data-message-id'), el.getAttribute('data-attachment-id'), el.getAttribute('data-mime-type'), el.getAttribute('data-filename'));
      };
      if (gmailBtn) gmailBtn.addEventListener('click', (e) => { e.stopPropagation(); fetchGmail(gmailBtn); });
      attachEl.addEventListener('click', (e) => { if (e.target.closest('.gmail-fetch-btn')) return; if (gmailBtn) fetchGmail(gmailBtn); });
    } else if (url) {
      attachEl.addEventListener('click', (e) => { if (e.target.closest('a')) return; isPreviewable ? showAttachmentPreview(url, type, name) : window.open(url, '_blank'); });
    }

    return attachEl;
  }

  // ============================================
  // showAttachmentPreview — Full-screen modal for images/videos/PDFs
  // ============================================
  function showAttachmentPreview(url, type, name) {
    document.querySelectorAll('.attachment-preview-modal').forEach(m => m.remove());
    const lt = (type || '').toLowerCase(), ln = (name || '').toLowerCase(), lu = (url || '').toLowerCase();
    let contentType = 'unknown';
    if (lt.includes('image') || lt === 'jpeg' || lt === 'png' || lt === 'gif' || ln.endsWith('.png') || ln.endsWith('.jpg') || ln.endsWith('.jpeg') || ln.endsWith('.gif') || ln.endsWith('.webp') || lu.includes('.png') || lu.includes('.jpg') || lu.includes('.jpeg') || lu.startsWith('data:image')) contentType = 'image';
    else if (lt.includes('video') || lt.includes('movie') || ln.endsWith('.mp4') || ln.endsWith('.mov') || ln.endsWith('.webm') || lu.includes('.mp4') || lu.includes('.mov')) contentType = 'video';
    else if (lt.includes('pdf') || ln.endsWith('.pdf') || lu.includes('.pdf')) contentType = 'pdf';

    if (contentType === 'unknown') { window.open(url, '_blank'); return; }

    const modal = document.createElement('div');
    modal.className = 'attachment-preview-modal';
    modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.9);z-index:100000;display:flex;flex-direction:column;animation:fadeIn 0.2s ease-out;';

    const header = document.createElement('div');
    header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:16px 20px;background:rgba(0,0,0,0.5);flex-shrink:0;';
    header.innerHTML = `
      <div style="display:flex;align-items:center;gap:12px;color:white;min-width:0;flex:1;">
        <span style="font-size:20px;">${contentType === 'image' ? '🖼️' : contentType === 'video' ? '🎬' : '📄'}</span>
        <span style="font-weight:600;font-size:14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Preview')}">${escapeHtml(name || 'Preview')}</span>
      </div>
      <div style="display:flex;align-items:center;gap:8px;flex-shrink:0;">
        <a href="${escapeHtml(url)}" target="_blank" style="display:flex;align-items:center;gap:6px;padding:8px 16px;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);border-radius:8px;color:white;text-decoration:none;font-size:13px;font-weight:500;"><span>↗</span> Open in new tab</a>
        <button class="preview-close-btn" style="display:flex;align-items:center;justify-content:center;width:40px;height:40px;background:rgba(231,76,60,0.8);border:none;border-radius:8px;color:white;font-size:20px;cursor:pointer;">×</button>
      </div>`;

    const content = document.createElement('div');
    content.style.cssText = 'flex:1;display:flex;align-items:center;justify-content:center;padding:20px;overflow:auto;min-height:0;';
    if (contentType === 'image') content.innerHTML = `<img src="${escapeHtml(url)}" alt="${escapeHtml(name || 'Image')}" style="max-width:100%;max-height:100%;object-fit:contain;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.3);" />`;
    else if (contentType === 'video') content.innerHTML = `<video controls autoplay style="max-width:100%;max-height:100%;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.3);"><source src="${escapeHtml(url)}" type="video/mp4">Your browser does not support video playback.</video>`;
    else if (contentType === 'pdf') {
      const gv = `https://docs.google.com/viewer?url=${encodeURIComponent(url)}&embedded=true`;
      content.innerHTML = `<div style="width:100%;height:100%;display:flex;flex-direction:column;gap:12px;"><object data="${escapeHtml(url)}" type="application/pdf" style="width:100%;height:100%;border-radius:8px;background:white;"><iframe src="${gv}" style="width:100%;height:100%;border:none;border-radius:8px;background:white;"></iframe></object></div>`;
      content.style.padding = '20px';
    }

    modal.appendChild(header);
    modal.appendChild(content);
    document.body.appendChild(modal);

    const closeModal = () => { modal.style.animation = 'fadeOut 0.2s ease-out'; setTimeout(() => modal.remove(), 180); };
    modal.querySelector('.preview-close-btn').addEventListener('click', (e) => { e.stopPropagation(); closeModal(); });
    modal.addEventListener('click', (e) => { e.stopPropagation(); if (e.target === modal) closeModal(); });
    content.addEventListener('click', (e) => { e.stopPropagation(); if (e.target === content) closeModal(); });
    const escH = (e) => { if (e.key === 'Escape') { closeModal(); document.removeEventListener('keydown', escH); } };
    document.addEventListener('keydown', escH);
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleMessageFormat = {
    formatMessageContent,
    sanitizeHtml,
    isComplexEmailHtml,
    stripEmailQuotedContent,
    renderEmailInIframe,
    fetchGmailAttachment,
    renderTranscriptAttachment,
    showAttachmentPreview,
  };

})();
